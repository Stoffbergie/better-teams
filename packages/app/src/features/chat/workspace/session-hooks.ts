import { useTeamsAccountContext } from "@better-teams/app/providers/TeamsAccountProvider";
import { teamsConversationService } from "@better-teams/app/services/teams/conversations";
import { teamsPresenceService } from "@better-teams/app/services/teams/presence";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import { teamsSessionService } from "@better-teams/app/services/teams/session";
import type {
  Conversation,
  TeamsSessionInfo,
} from "@better-teams/core/teams/types";
import { useQuery } from "@tanstack/react-query";
import {
  useDocumentVisibility,
  useResumeCooldown,
} from "./document-visibility";

const SELF_AVAILABILITY_HEARTBEAT_MS = 60_000;
const EMPTY_CONVERSATIONS: Conversation[] = [];

export function useActiveTeamsAccount() {
  return useTeamsAccountContext();
}

export function useTeamsSession(): {
  tenantId?: string;
  session?: TeamsSessionInfo;
  isPending: boolean;
  isFetching: boolean;
  isError: boolean;
  error: Error | null;
  refetch: () => Promise<unknown>;
} {
  const { activeTenantId, activeSession } = useTeamsAccountContext();
  const { data, error, isError, isFetching, isPending, refetch } = useQuery({
    queryKey: teamsKeys.session(activeTenantId),
    queryFn: () => teamsSessionService.initialize(activeTenantId),
    initialData: activeSession,
    initialDataUpdatedAt: activeSession ? 0 : undefined,
    enabled: true,
    staleTime: 30_000,
    gcTime: Number.POSITIVE_INFINITY,
  });

  return {
    tenantId: activeTenantId,
    session: data ?? activeSession,
    isPending,
    isFetching,
    isError,
    error: error instanceof Error ? error : null,
    refetch,
  };
}

export function useTeamsConversations(liveSessionReady: boolean): {
  tenantId?: string;
  conversations: Conversation[];
  isPending: boolean;
  isFetching: boolean;
  isError: boolean;
  isSuccess: boolean;
  refetch: () => Promise<unknown>;
} {
  const { activeTenantId } = useTeamsAccountContext();
  const documentVisible = useDocumentVisibility();
  const resumeReady = useResumeCooldown();

  const { data, isError, isFetching, isPending, isSuccess, refetch } = useQuery(
    {
      queryKey: teamsKeys.conversations(activeTenantId),
      queryFn: () => teamsConversationService.list(activeTenantId, 100),
      enabled: liveSessionReady,
      staleTime: 30_000,
      refetchInterval: () =>
        liveSessionReady && documentVisible && resumeReady ? 30_000 : false,
      refetchIntervalInBackground: false,
    },
  );

  return {
    tenantId: activeTenantId,
    conversations: data ?? EMPTY_CONVERSATIONS,
    isPending,
    isFetching,
    isError,
    isSuccess,
    refetch,
  };
}

export function useMaintainTeamsAvailability(enabled: boolean) {
  const { activeTenantId } = useTeamsAccountContext();
  const documentVisible = useDocumentVisibility();
  const resumeReady = useResumeCooldown();

  useQuery({
    queryKey: teamsKeys.selfAvailability(activeTenantId),
    queryFn: async () => {
      await teamsPresenceService.setSelfAvailability(activeTenantId);
      return Date.now();
    },
    enabled: enabled && documentVisible && resumeReady,
    staleTime: SELF_AVAILABILITY_HEARTBEAT_MS - 5_000,
    gcTime: 5 * 60_000,
    refetchInterval: SELF_AVAILABILITY_HEARTBEAT_MS,
    refetchIntervalInBackground: false,
    refetchOnWindowFocus: false,
    retry: 1,
  });
}
