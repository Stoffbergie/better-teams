import { useTeamsAccountContext } from "@better-teams/app/providers/TeamsAccountProvider";
import { getCachedPresence } from "@better-teams/app/services/desktop/runtime";
import { teamsPresenceService } from "@better-teams/app/services/teams/presence";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import {
  canonAvatarMri,
  collectProfileAvatarMris,
} from "@better-teams/core/teams/profile/avatars";
import type {
  Conversation,
  Message,
  PresenceInfo,
} from "@better-teams/core/teams/types";
import { useQuery } from "@tanstack/react-query";
import { useDeferredValue, useMemo } from "react";
import {
  useDocumentVisibility,
  useResumeCooldown,
} from "../workspace/document-visibility";

const PRESENCE_BATCH_SIZE = 50;
const EMPTY_MESSAGES: Message[] = [];
const EMPTY_PRESENCE_BY_MRI: Record<string, PresenceInfo> = {};

function uniqueMris(mris: string[]): string[] {
  const unique = new Map<string, string>();
  for (const mri of mris) {
    const trimmed = mri.trim();
    if (!trimmed) continue;
    unique.set(canonAvatarMri(trimmed), trimmed);
  }
  return [...unique.values()];
}

async function fetchPresence(
  tenantId: string | undefined,
  mris: string[],
): Promise<Record<string, PresenceInfo>> {
  const cachedPresence = await getCachedPresence(mris);
  const presence: Record<string, PresenceInfo> = Object.fromEntries(
    Object.entries(cachedPresence).map(([mri, info]) => [
      canonAvatarMri(mri),
      info,
    ]),
  );
  const missingMris = mris.filter((mri) => !(canonAvatarMri(mri) in presence));
  if (missingMris.length === 0) {
    return presence;
  }

  for (
    let index = 0;
    index < missingMris.length;
    index += PRESENCE_BATCH_SIZE
  ) {
    const batch = missingMris.slice(index, index + PRESENCE_BATCH_SIZE);
    try {
      const batchPresence = await teamsPresenceService.getPresence(
        tenantId,
        batch,
      );
      for (const [mri, info] of Object.entries(batchPresence)) {
        presence[canonAvatarMri(mri)] = info;
      }
    } catch (error) {
      if (Object.keys(presence).length > 0) {
        continue;
      }
      throw error;
    }
  }

  return presence;
}

export function useTeamsPresence(args: {
  conversations: Conversation[];
  messages?: Message[];
  selfSkypeId?: string;
}) {
  const { activeTenantId } = useTeamsAccountContext();
  const documentVisible = useDocumentVisibility();
  const resumeReady = useResumeCooldown();
  const deferredConversations = useDeferredValue(args.conversations);
  const deferredMessages = useDeferredValue(args.messages ?? EMPTY_MESSAGES);
  const presenceMris = useMemo(
    () =>
      uniqueMris(
        collectProfileAvatarMris({
          conversations: deferredConversations,
          messages: deferredMessages,
          selfSkypeId: args.selfSkypeId,
        }),
      ),
    [deferredConversations, deferredMessages, args.selfSkypeId],
  );
  const signature = useMemo(
    () => [...presenceMris].sort().join("\x1f"),
    [presenceMris],
  );

  const { data } = useQuery({
    queryKey: teamsKeys.presence(activeTenantId, signature),
    queryFn: () => fetchPresence(activeTenantId, presenceMris),
    enabled: documentVisible && resumeReady && presenceMris.length > 0,
    staleTime: 30_000,
    gcTime: 5 * 60_000,
    placeholderData: (previousData) => previousData,
    refetchInterval: () => (documentVisible && resumeReady ? 60_000 : false),
    refetchIntervalInBackground: false,
    refetchOnWindowFocus: false,
    refetchOnReconnect: false,
    retry: 1,
  });

  return data ?? EMPTY_PRESENCE_BY_MRI;
}
