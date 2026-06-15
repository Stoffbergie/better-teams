import { useTeamsAccountContext } from "@better-teams/app/providers/TeamsAccountProvider";
import { teamsProfileService } from "@better-teams/app/services/teams/profile";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import {
  canonAvatarMri,
  collectProfileAvatarMris,
} from "@better-teams/core/teams/profile/avatars";
import { TeamsProfilePresentationSchema } from "@better-teams/core/teams/schemas";
import type {
  Conversation,
  Message,
  TeamsProfilePresentation,
} from "@better-teams/core/teams/types";
import { useQuery } from "@tanstack/react-query";
import { useDeferredValue, useMemo } from "react";

const PRIORITY_AVATAR_CONVERSATIONS = 12;
const PROFILE_QUERY_GC_MS = 5 * 60_000;
const EMPTY_MESSAGES: Message[] = [];

export type TeamsProfilePresentationState = TeamsProfilePresentation & {
  avatarFallbackReady: boolean;
};

const EMPTY_PROFILE_PRESENTATION: TeamsProfilePresentation = {
  avatarThumbs: {},
  avatarFull: {},
  displayNames: {},
  emails: {},
  jobTitles: {},
  departments: {},
  companyNames: {},
  tenantNames: {},
  locations: {},
};
const profilePersonCacheByTenant = new Map<
  string,
  Map<string, TeamsProfilePresentation>
>();

function normalizeProfilePresentation(data: unknown): TeamsProfilePresentation {
  const parsed = TeamsProfilePresentationSchema.safeParse(data);
  if (parsed.success) return parsed.data;
  const raw =
    data && typeof data === "object" ? (data as Record<string, unknown>) : {};
  const record = (value: unknown): Record<string, string> => {
    if (!value || typeof value !== "object" || Array.isArray(value)) return {};
    return Object.fromEntries(
      Object.entries(value).filter(
        (entry): entry is [string, string] => typeof entry[1] === "string",
      ),
    );
  };
  return {
    avatarThumbs: record(raw.avatarThumbs ?? raw.avatars),
    avatarFull: record(raw.avatarFull ?? raw.avatars),
    displayNames: record(raw.displayNames),
    emails: record(raw.emails),
    jobTitles: record(raw.jobTitles),
    departments: record(raw.departments),
    companyNames: record(raw.companyNames),
    tenantNames: record(raw.tenantNames),
    locations: record(raw.locations),
  };
}

async function fetchProfiles(
  tenantId: string | undefined,
  mris: string[],
): Promise<TeamsProfilePresentation> {
  if (mris.length === 0) return normalizeProfilePresentation(null);
  return teamsProfileService.fetchPresentation(tenantId, mris);
}

function mergeProfilePresentations(
  presentations: unknown[],
): TeamsProfilePresentation {
  return presentations.reduce<TeamsProfilePresentation>((merged, data) => {
    const next = normalizeProfilePresentation(data);
    return {
      avatarThumbs: { ...merged.avatarThumbs, ...next.avatarThumbs },
      avatarFull: { ...merged.avatarFull, ...next.avatarFull },
      displayNames: { ...merged.displayNames, ...next.displayNames },
      emails: { ...merged.emails, ...next.emails },
      jobTitles: { ...merged.jobTitles, ...next.jobTitles },
      departments: { ...merged.departments, ...next.departments },
      companyNames: { ...merged.companyNames, ...next.companyNames },
      tenantNames: { ...merged.tenantNames, ...next.tenantNames },
      locations: { ...merged.locations, ...next.locations },
    };
  }, EMPTY_PROFILE_PRESENTATION);
}

function profilePresentationForMri(
  presentation: TeamsProfilePresentation,
  mri: string,
): TeamsProfilePresentation {
  const key = canonAvatarMri(mri);
  const pick = (record: Record<string, string>) =>
    record[key] ? { [key]: record[key] } : {};
  return {
    avatarThumbs: pick(presentation.avatarThumbs),
    avatarFull: pick(presentation.avatarFull),
    displayNames: pick(presentation.displayNames),
    emails: pick(presentation.emails),
    jobTitles: pick(presentation.jobTitles),
    departments: pick(presentation.departments),
    companyNames: pick(presentation.companyNames),
    tenantNames: pick(presentation.tenantNames),
    locations: pick(presentation.locations),
  };
}

function tenantProfilePersonCache(tenantId: string | null | undefined) {
  const tenantKey = tenantId ?? "__default__";
  let cache = profilePersonCacheByTenant.get(tenantKey);
  if (!cache) {
    cache = new Map();
    profilePersonCacheByTenant.set(tenantKey, cache);
  }
  return cache;
}

function readProfilePersonCache(
  tenantId: string | null | undefined,
  mris: string[],
): {
  presentation: TeamsProfilePresentation;
  missingMris: string[];
} {
  const cache = tenantProfilePersonCache(tenantId);
  const cached: TeamsProfilePresentation[] = [];
  const missingMris: string[] = [];
  const seen = new Set<string>();

  for (const mri of mris) {
    const normalizedMri = canonAvatarMri(mri);
    if (!normalizedMri || seen.has(normalizedMri)) continue;
    seen.add(normalizedMri);
    const data = cache.get(normalizedMri);
    if (data === undefined) {
      missingMris.push(mri);
    } else {
      cached.push(data);
    }
  }

  return {
    presentation: mergeProfilePresentations(cached),
    missingMris,
  };
}

function seedProfilePersonCache(
  tenantId: string | null | undefined,
  mris: string[],
  data: unknown,
) {
  const cache = tenantProfilePersonCache(tenantId);
  const presentation = normalizeProfilePresentation(data);
  const seen = new Set<string>();

  for (const mri of mris) {
    const normalizedMri = canonAvatarMri(mri);
    if (!normalizedMri || seen.has(normalizedMri)) continue;
    seen.add(normalizedMri);
    const personPresentation = profilePresentationForMri(
      presentation,
      normalizedMri,
    );
    cache.set(
      normalizedMri,
      mergeProfilePresentations([cache.get(normalizedMri), personPresentation]),
    );
  }
}

async function fetchProfilesWithPersonCache({
  tenantId,
  mris,
}: {
  tenantId: string | null | undefined;
  mris: string[];
}): Promise<TeamsProfilePresentation> {
  const cached = readProfilePersonCache(tenantId, mris);
  if (cached.missingMris.length === 0) return cached.presentation;

  const fetched = await fetchProfiles(
    tenantId ?? undefined,
    cached.missingMris,
  );
  seedProfilePersonCache(tenantId, cached.missingMris, fetched);
  return mergeProfilePresentations([cached.presentation, fetched]);
}

export function useTeamsProfilePresentation(args: {
  conversations: Conversation[];
  messages?: Message[];
  selfSkypeId?: string;
}): TeamsProfilePresentationState {
  const { activeTenantId } = useTeamsAccountContext();
  const deferredConversations = useDeferredValue(args.conversations);
  const deferredMessages = useDeferredValue(args.messages ?? EMPTY_MESSAGES);
  const profileMris = useMemo(
    () =>
      collectProfileAvatarMris({
        conversations: deferredConversations,
        messages: deferredMessages,
        selfSkypeId: args.selfSkypeId,
      }),
    [deferredConversations, deferredMessages, args.selfSkypeId],
  );
  const priorityProfileMris = useMemo(
    () => profileMris.slice(0, PRIORITY_AVATAR_CONVERSATIONS),
    [profileMris],
  );
  const profileMriSignature = useMemo(
    () => [...profileMris].sort().join("\x1f"),
    [profileMris],
  );
  const prioritySignature = useMemo(
    () => [...priorityProfileMris].sort().join("\x1f"),
    [priorityProfileMris],
  );
  const priorityPersonCache = useMemo(
    () => readProfilePersonCache(activeTenantId, priorityProfileMris),
    [activeTenantId, priorityProfileMris],
  );
  const profilePersonCache = useMemo(
    () => readProfilePersonCache(activeTenantId, profileMris),
    [activeTenantId, profileMris],
  );

  const {
    data: priorityAvatarData,
    isFetching: priorityAvatarFetching,
    isPending: priorityAvatarPending,
    isSuccess: priorityAvatarSuccess,
  } = useQuery({
    queryKey: teamsKeys.profileAvatars(activeTenantId, prioritySignature),
    queryFn: () =>
      fetchProfilesWithPersonCache({
        tenantId: activeTenantId,
        mris: priorityProfileMris,
      }),
    enabled: Boolean(activeTenantId) && priorityProfileMris.length > 0,
    staleTime: 3_600_000,
    gcTime: PROFILE_QUERY_GC_MS,
    placeholderData: (previousData) =>
      mergeProfilePresentations([
        priorityPersonCache.presentation,
        previousData,
      ]),
    retry: 2,
    retryDelay: (attempt) => Math.min(2000 * 2 ** attempt, 12_000),
  });

  const {
    data: backgroundAvatarData,
    isFetching: backgroundAvatarFetching,
    isPending: backgroundAvatarPending,
  } = useQuery({
    queryKey: teamsKeys.profileAvatars(activeTenantId, profileMriSignature),
    queryFn: () =>
      fetchProfilesWithPersonCache({
        tenantId: activeTenantId,
        mris: profileMris,
      }),
    enabled:
      Boolean(activeTenantId) &&
      profileMris.length > 0 &&
      profileMriSignature !== prioritySignature &&
      priorityAvatarSuccess,
    staleTime: 3_600_000,
    gcTime: PROFILE_QUERY_GC_MS,
    placeholderData: () =>
      mergeProfilePresentations([
        profilePersonCache.presentation,
        priorityAvatarData,
      ]),
    retry: 2,
    retryDelay: (attempt) => Math.min(2000 * 2 ** attempt, 12_000),
  });

  const profilePresentationData = useMemo(
    () =>
      mergeProfilePresentations([
        profilePersonCache.presentation,
        priorityAvatarData,
        backgroundAvatarData,
      ]),
    [backgroundAvatarData, priorityAvatarData, profilePersonCache.presentation],
  );
  const isPriorityAvatarLoading =
    Boolean(activeTenantId) &&
    priorityProfileMris.length > 0 &&
    (priorityAvatarPending || priorityAvatarFetching);
  const isBackgroundAvatarLoading =
    Boolean(activeTenantId) &&
    profileMris.length > 0 &&
    profileMriSignature !== prioritySignature &&
    priorityAvatarSuccess &&
    (backgroundAvatarPending || backgroundAvatarFetching);
  const profilePersonCacheComplete =
    profilePersonCache.missingMris.length === 0;

  const profilePresentation = useMemo(
    () => normalizeProfilePresentation(profilePresentationData),
    [profilePresentationData],
  );

  return useMemo(
    () => ({
      ...profilePresentation,
      avatarFallbackReady:
        profilePersonCacheComplete ||
        (!isPriorityAvatarLoading && !isBackgroundAvatarLoading),
    }),
    [
      isBackgroundAvatarLoading,
      isPriorityAvatarLoading,
      profilePersonCacheComplete,
      profilePresentation,
    ],
  );
}
