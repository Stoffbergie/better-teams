import { teamsConversationService } from "@better-teams/app/services/teams/conversations";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import { buildSharedConversationsByMri } from "@better-teams/core/teams/conversation/shared-conversation-index";
import type { ConversationMember } from "@better-teams/core/teams/types";
import { type QueryClient, useQuery } from "@tanstack/react-query";
import { useMemo } from "react";
import type { SidebarConversationItem } from "../thread/types";

export function useSharedConversationLookup({
  activeTenantId,
  profileSidebarMri,
  allSidebarItems,
  sidebarItemById,
  sidebarDisplayNameByMri,
  displayNameByMri,
  emailByMri,
  queryClient,
}: {
  activeTenantId?: string | null;
  profileSidebarMri: string | null;
  allSidebarItems: SidebarConversationItem[];
  sidebarItemById: Record<string, SidebarConversationItem>;
  sidebarDisplayNameByMri: Record<string, string>;
  displayNameByMri: Record<string, string>;
  emailByMri: Record<string, string>;
  queryClient: QueryClient;
}) {
  const sharedConversationCandidateIds = useMemo(
    () =>
      allSidebarItems
        .filter((item) => item.kind !== "dm")
        .map((item) => item.id)
        .sort(),
    [allSidebarItems],
  );
  const {
    data: sharedConversationDetails,
    isPending: sharedConversationDetailsPending,
  } = useQuery({
    queryKey: [
      "teams",
      "shared-thread-members",
      activeTenantId ?? "__default__",
      profileSidebarMri ?? "__none__",
      sharedConversationCandidateIds.join("\x1f"),
    ],
    queryFn: async () => {
      const byConversationId: Record<string, ConversationMember[]> = {};
      const missingConversationIds: string[] = [];

      for (const conversationId of sharedConversationCandidateIds) {
        const cachedMembers = queryClient.getQueryData<ConversationMember[]>(
          teamsKeys.threadMembers(activeTenantId, conversationId),
        );
        if (cachedMembers && cachedMembers.length > 0) {
          byConversationId[conversationId] = cachedMembers;
          continue;
        }

        const sidebarMembers =
          sidebarItemById[conversationId]?.conversation.members;
        if (sidebarMembers && sidebarMembers.length > 0) {
          byConversationId[conversationId] = sidebarMembers;
          queryClient.setQueryData(
            teamsKeys.threadMembers(activeTenantId, conversationId),
            sidebarMembers,
          );
          continue;
        }

        missingConversationIds.push(conversationId);
      }

      if (missingConversationIds.length === 0) {
        return byConversationId;
      }

      const concurrency = 6;
      let index = 0;

      await Promise.all(
        Array.from(
          { length: Math.min(concurrency, missingConversationIds.length) },
          async () => {
            while (index < missingConversationIds.length) {
              const conversationId = missingConversationIds[index];
              index += 1;
              if (!conversationId) continue;
              try {
                const members = await teamsConversationService.getMembers(
                  activeTenantId,
                  conversationId,
                );
                byConversationId[conversationId] = members;
                queryClient.setQueryData(
                  teamsKeys.threadMembers(activeTenantId, conversationId),
                  members,
                );
              } catch {}
            }
          },
        ),
      );

      return byConversationId;
    },
    enabled:
      Boolean(profileSidebarMri) && sharedConversationCandidateIds.length > 0,
    staleTime: 5 * 60_000,
  });
  const detailedSharedConversationById = sharedConversationDetails ?? {};

  const sharedConversationsByMri = useMemo(
    () =>
      buildSharedConversationsByMri(
        allSidebarItems,
        detailedSharedConversationById,
        { ...sidebarDisplayNameByMri, ...displayNameByMri },
        emailByMri,
      ),
    [
      allSidebarItems,
      detailedSharedConversationById,
      sidebarDisplayNameByMri,
      displayNameByMri,
      emailByMri,
    ],
  );

  return {
    sharedConversationsByMri,
    loading:
      Boolean(profileSidebarMri) &&
      sharedConversationCandidateIds.length > 0 &&
      sharedConversationDetailsPending,
  };
}
