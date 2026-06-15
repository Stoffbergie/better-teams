import {
  useTeamsPresence,
  useTeamsProfilePresentation,
} from "@better-teams/app/features/chat/workspace/teams-hooks";
import {
  countDomNodes,
  isPerfEnabled,
  measurePerfAsync,
  updatePerfSnapshot,
} from "@better-teams/app/platform/perf";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import { teamsThreadService } from "@better-teams/app/services/teams/thread";
import {
  formatDetailedTimestamp,
  formatMessageTime,
  formatThreadDayDividerLabel,
  gapBetweenMessages,
  isEditedMessage,
  isSelfMessage,
  type MessageInlinePart,
  messageReadStatus,
  messageReadTimestamp,
  messageRichPartsForDisplay,
  messageTimestamp,
  parseConsumptionHorizon,
} from "@better-teams/core/chat";
import { canonAvatarMri } from "@better-teams/core/teams/profile/avatars";
import {
  sortMessagesByTimestamp,
  type ThreadQueryData,
  threadQueryDataFromResponse,
} from "@better-teams/core/teams/thread";
import type {
  Conversation,
  ConversationMember,
} from "@better-teams/core/teams/types";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import {
  forwardRef,
  useCallback,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
} from "react";
import { MessageRow } from "../message/MessageRow";
import type { ProfileData } from "../profile/ProfileCard";
import { useThreadScroll } from "./thread-scroll";
import { useThreadSearch } from "./thread-search";
import { type DisplayMessage, type MessageBlock, THREAD_PAGE } from "./types";

export {
  captureScrollRestoreAnchor,
  olderPrefetchThresholdForVelocity,
  restoreScrollRestoreAnchor,
  shouldPrefetchOlderMessages,
} from "./thread-scroll";

type MemberReadReceipt = {
  mri: string;
  sequenceId: number;
  timestamp: number;
};

function selfMriFromSkypeId(skypeId?: string): string | null {
  if (!skypeId) return null;
  const trimmed = skypeId.trim();
  if (!trimmed) return null;
  return canonAvatarMri(trimmed.startsWith("8:") ? trimmed : `8:${trimmed}`);
}

function memberName(
  member: ConversationMember,
  displayNameByMri: Record<string, string>,
): string {
  const mri = canonAvatarMri(member.id);
  return (
    displayNameByMri[mri] ||
    member.displayName?.trim() ||
    member.friendlyName?.trim() ||
    member.userPrincipalName?.trim() ||
    mri
  );
}

export type ThreadViewHandle = {
  submitSearch: (query: string) => void;
};

type ThreadViewProps = {
  tenantId?: string | null;
  conversationId: string;
  conversationKind: "dm" | "group" | "meeting";
  liveSessionReady: boolean;
  autoFocus?: boolean;
  searchQuery: string;
  consumptionHorizon?: string;
  onSearchResultCountChange?: (resultCount: number) => void;
  selfSkypeId?: string;
  selfDisplayName?: string;
  avatarByMri: Record<string, string>;
  avatarFullByMri: Record<string, string>;
  avatarFallbackReady: boolean;
  displayNameByMri: Record<string, string>;
  emailByMri: Record<string, string>;
  jobTitleByMri: Record<string, string>;
  departmentByMri: Record<string, string>;
  companyNameByMri: Record<string, string>;
  tenantNameByMri: Record<string, string>;
  locationByMri: Record<string, string>;
  sharedConversationsByMri: Record<
    string,
    NonNullable<ProfileData["sharedConversations"]>
  >;
  onOpenProfile?: (profile: ProfileData) => void;
};

export function profileMessageConversationId(
  conversationKind: "dm" | "group" | "meeting",
  _conversationId: string,
  sharedConversations: NonNullable<ProfileData["sharedConversations"]>,
): string | undefined {
  if (conversationKind === "dm") return undefined;
  return sharedConversations.find((conversation) => conversation.kind === "dm")
    ?.id;
}

export const ThreadView = forwardRef<ThreadViewHandle, ThreadViewProps>(
  (
    {
      tenantId,
      conversationId,
      conversationKind,
      liveSessionReady,
      autoFocus,
      searchQuery,
      consumptionHorizon,
      onSearchResultCountChange,
      selfSkypeId,
      selfDisplayName,
      avatarByMri,
      avatarFullByMri,
      avatarFallbackReady,
      displayNameByMri,
      emailByMri,
      jobTitleByMri,
      departmentByMri,
      companyNameByMri,
      tenantNameByMri,
      locationByMri,
      sharedConversationsByMri,
      onOpenProfile,
    },
    ref,
  ) => {
    const queryClient = useQueryClient();

    const perfEnabled = isPerfEnabled();

    const viewportRef = useRef<HTMLDivElement>(null);
    const topSentinelRef = useRef<HTMLLIElement>(null);

    const {
      data: threadQueryData,
      isError: threadQueryError,
      isPending: threadQueryPending,
      refetch: refetchThread,
    } = useQuery({
      queryKey: teamsKeys.thread(tenantId, conversationId),
      queryFn: async () => {
        return measurePerfAsync(
          "thread.fetchMessages",
          { conversationId, pageSize: THREAD_PAGE, tenantId: tenantId ?? null },
          async () => {
            const res = await teamsThreadService.getMessages(
              tenantId,
              conversationId,
              THREAD_PAGE,
              1,
            );
            return threadQueryDataFromResponse(res);
          },
        );
      },
      enabled: liveSessionReady,
      staleTime: 25_000,
    });
    const { data: threadMembersData } = useQuery({
      queryKey: teamsKeys.threadMembers(tenantId, conversationId),
      queryFn: async () =>
        measurePerfAsync(
          "thread.fetchMembers",
          { conversationId, tenantId: tenantId ?? null },
          () => teamsThreadService.getMembers(tenantId, conversationId),
        ),
      enabled: liveSessionReady,
      staleTime: 60_000,
    });
    const { data: consumptionHorizonsData } = useQuery({
      queryKey: teamsKeys.threadConsumptionHorizons(tenantId, conversationId),
      queryFn: async () => {
        return teamsThreadService.getMembersConsumptionHorizon(
          tenantId,
          conversationId,
        );
      },
      enabled: liveSessionReady && conversationKind !== "meeting",
      staleTime: 20_000,
      retry: 1,
    });

    const threadData = threadQueryData ?? null;
    const rawMessages = threadData?.messages ?? [];
    const threadMembers = threadMembersData ?? [];
    const threadLoading = threadQueryPending;
    const threadHasData = Boolean(threadData);
    const profileConversations = useMemo<Conversation[]>(
      () =>
        threadMembers.length > 0
          ? [
              {
                id: conversationId,
                members: threadMembers,
              } satisfies Conversation,
            ]
          : [],
      [conversationId, threadMembers],
    );
    const threadProfilePresentation = useTeamsProfilePresentation({
      conversations: profileConversations,
      messages: rawMessages,
      selfSkypeId,
    });
    const threadPresenceByMri = useTeamsPresence({
      conversations: profileConversations,
      messages: rawMessages,
      selfSkypeId,
    });
    const mergedAvatarByMri = useMemo(
      () => ({ ...avatarByMri, ...threadProfilePresentation.avatarThumbs }),
      [avatarByMri, threadProfilePresentation.avatarThumbs],
    );
    const mergedAvatarFullByMri = useMemo(
      () => ({ ...avatarFullByMri, ...threadProfilePresentation.avatarFull }),
      [avatarFullByMri, threadProfilePresentation.avatarFull],
    );
    const mergedAvatarFallbackReady =
      avatarFallbackReady && threadProfilePresentation.avatarFallbackReady;
    const mergedDisplayNameByMri = useMemo(
      () => ({
        ...displayNameByMri,
        ...threadProfilePresentation.displayNames,
      }),
      [displayNameByMri, threadProfilePresentation.displayNames],
    );
    const mergedEmailByMri = useMemo(
      () => ({ ...emailByMri, ...threadProfilePresentation.emails }),
      [emailByMri, threadProfilePresentation.emails],
    );
    const mergedJobTitleByMri = useMemo(
      () => ({
        ...jobTitleByMri,
        ...threadProfilePresentation.jobTitles,
      }),
      [jobTitleByMri, threadProfilePresentation.jobTitles],
    );
    const mergedDepartmentByMri = useMemo(
      () => ({
        ...departmentByMri,
        ...threadProfilePresentation.departments,
      }),
      [departmentByMri, threadProfilePresentation.departments],
    );
    const mergedCompanyNameByMri = useMemo(
      () => ({
        ...companyNameByMri,
        ...threadProfilePresentation.companyNames,
      }),
      [companyNameByMri, threadProfilePresentation.companyNames],
    );
    const mergedTenantNameByMri = useMemo(
      () => ({
        ...tenantNameByMri,
        ...threadProfilePresentation.tenantNames,
      }),
      [tenantNameByMri, threadProfilePresentation.tenantNames],
    );
    const mergedLocationByMri = useMemo(
      () => ({
        ...locationByMri,
        ...threadProfilePresentation.locations,
      }),
      [locationByMri, threadProfilePresentation.locations],
    );
    const selfMri = useMemo(
      () => selfMriFromSkypeId(selfSkypeId),
      [selfSkypeId],
    );
    const selfMessageDisplayName =
      (selfMri ? mergedDisplayNameByMri[selfMri]?.trim() : "") ||
      selfDisplayName?.trim() ||
      "You";
    const fallbackPeerHorizons = useMemo(() => {
      const h = parseConsumptionHorizon(consumptionHorizon);
      return h ? [h] : [];
    }, [consumptionHorizon]);
    const memberHorizons = useMemo<MemberReadReceipt[]>(() => {
      return (consumptionHorizonsData?.consumptionhorizons ?? [])
        .map((entry) => {
          const parsed = parseConsumptionHorizon(entry.consumptionhorizon);
          if (!parsed) return null;
          return {
            mri: canonAvatarMri(entry.id),
            sequenceId: parsed.sequenceId,
            timestamp: parsed.timestamp,
          } satisfies MemberReadReceipt;
        })
        .filter((entry): entry is MemberReadReceipt => entry != null);
    }, [consumptionHorizonsData?.consumptionhorizons]);
    const peerHorizons = useMemo(() => {
      if (memberHorizons.length > 0) {
        return memberHorizons
          .filter((entry) => entry.mri !== selfMri)
          .map((entry) => ({
            sequenceId: entry.sequenceId,
            timestamp: entry.timestamp,
            messageId: entry.mri,
          }));
      }
      return fallbackPeerHorizons;
    }, [fallbackPeerHorizons, memberHorizons, selfMri]);
    const receiptParticipants = useMemo(
      () =>
        threadMembers.filter((member) => canonAvatarMri(member.id) !== selfMri),
      [selfMri, threadMembers],
    );
    const receiptParticipantByMri = useMemo(
      () =>
        Object.fromEntries(
          receiptParticipants.map((member) => [
            canonAvatarMri(member.id),
            member,
          ]),
        ),
      [receiptParticipants],
    );

    const threadDisplayState = useMemo(() => {
      const displayMessages = rawMessages.flatMap((message) => {
        const parts = messageRichPartsForDisplay(message);
        const bodyText = parts?.body.map((part) => part.text).join("") ?? "";
        const quoteText = parts?.quote?.map((part) => part.text).join("") ?? "";
        const attachmentTitles =
          parts?.attachments.map((attachment) => attachment.title).join(" ") ??
          "";
        if (
          !parts ||
          (!bodyText.trim() &&
            !quoteText.trim() &&
            parts.attachments.length === 0 &&
            !message.deleted)
        ) {
          return [];
        }
        const self = isSelfMessage(message.from, selfSkypeId);
        const messageSequenceId =
          message.sequenceId ?? (message.id ? Number(message.id) : Number.NaN);
        const seenBy =
          self &&
          conversationKind === "group" &&
          Number.isFinite(messageSequenceId)
            ? memberHorizons
                .filter(
                  (entry) =>
                    entry.mri !== selfMri &&
                    entry.sequenceId >= messageSequenceId,
                )
                .sort((left, right) => right.timestamp - left.timestamp)
                .map((entry) => {
                  const member = receiptParticipantByMri[entry.mri];
                  return {
                    mri: entry.mri,
                    name: member
                      ? memberName(member, mergedDisplayNameByMri)
                      : mergedDisplayNameByMri[entry.mri] || entry.mri,
                    readAt: formatDetailedTimestamp(
                      new Date(entry.timestamp).toISOString(),
                    ),
                  };
                })
            : [];
        const seenMris = new Set(seenBy.map((entry) => entry.mri));
        const unseenBy =
          self && conversationKind === "group"
            ? receiptParticipants
                .filter((member) => !seenMris.has(canonAvatarMri(member.id)))
                .map((member) => ({
                  mri: canonAvatarMri(member.id),
                  name: memberName(member, mergedDisplayNameByMri),
                }))
            : [];
        return [
          {
            message,
            parts,
            displayName: self
              ? selfMessageDisplayName
              : message.senderDisplayName?.trim() ||
                message.imdisplayname?.trim() ||
                "Unknown",
            time: formatMessageTime(messageTimestamp(message)),
            self,
            deleted: Boolean(message.deleted),
            edited: isEditedMessage(message),
            readStatus: self
              ? messageReadStatus(message, peerHorizons)
              : undefined,
            sentAt: self
              ? formatDetailedTimestamp(messageTimestamp(message))
              : "",
            readAt: self
              ? formatDetailedTimestamp(
                  messageReadTimestamp(message, peerHorizons),
                )
              : "",
            receiptScope:
              conversationKind === "dm" ? ("dm" as const) : ("group" as const),
            receiptSeenBy: seenBy,
            receiptUnseenBy: unseenBy,
            bodyPreview: [bodyText, quoteText, attachmentTitles]
              .join(" ")
              .replace(/\s+/g, " ")
              .trim(),
            searchText: [
              self
                ? selfMessageDisplayName
                : message.senderDisplayName?.trim() ||
                  message.imdisplayname?.trim() ||
                  "Unknown",
              bodyText,
              quoteText,
              attachmentTitles,
            ]
              .join(" ")
              .toLowerCase(),
          },
        ];
      });
      const messageBlocks: MessageBlock[] = [];
      let lastDay = "";
      let previous: DisplayMessage | undefined;
      for (let i = 0; i < displayMessages.length; i++) {
        const entry = displayMessages[i];
        const ts = messageTimestamp(entry.message);
        const day = ts ? formatThreadDayDividerLabel(ts) : "";
        if (day && day !== lastDay) {
          lastDay = day;
          messageBlocks.push({
            kind: "day",
            label: day,
            key: `day-${i}-${day}`,
          });
        }
        const showMeta =
          !previous ||
          previous.message.from !== entry.message.from ||
          gapBetweenMessages(previous.message, entry.message) > 5 * 60 * 1000;
        messageBlocks.push({
          kind: "msg",
          entry,
          messageIndex: i,
          showMeta,
          key: entry.message.id,
        });
        previous = entry;
      }
      return { displayMessages, messageBlocks };
    }, [
      conversationKind,
      memberHorizons,
      mergedDisplayNameByMri,
      rawMessages,
      receiptParticipantByMri,
      receiptParticipants,
      selfSkypeId,
      selfMessageDisplayName,
      selfMri,
      peerHorizons,
    ]);
    const displayMessages = threadDisplayState.displayMessages;
    const messageBlocks = threadDisplayState.messageBlocks;
    const loadedMessageCount = rawMessages.length;

    const mentionProfileForPart = useCallback(
      (part: MessageInlinePart): ProfileData | null => {
        if (part.kind !== "mention" || !part.mentionedMri) return null;
        const mri = canonAvatarMri(part.mentionedMri);
        const messageConversationId = profileMessageConversationId(
          conversationKind,
          conversationId,
          sharedConversationsByMri[mri] ?? [],
        );
        return {
          mri,
          displayName:
            mergedDisplayNameByMri[mri] ||
            part.mentionedDisplayName ||
            part.text.replace(/^@/, ""),
          avatarThumbSrc: mergedAvatarByMri[mri],
          avatarFullSrc: mergedAvatarFullByMri[mri] ?? mergedAvatarByMri[mri],
          avatarFallbackReady: mergedAvatarFallbackReady,
          email: mergedEmailByMri[mri],
          jobTitle: mergedJobTitleByMri[mri],
          department: mergedDepartmentByMri[mri],
          companyName: mergedCompanyNameByMri[mri],
          tenantName: mergedTenantNameByMri[mri],
          location: mergedLocationByMri[mri],
          presence: threadPresenceByMri[mri],
          onOpenConversation: (targetConversationId: string) => {
            queryClient.setQueryData<string | null>(
              ["open-conversation-request"],
              targetConversationId,
            );
          },
          onMessage: messageConversationId
            ? () => {
                queryClient.setQueryData<string | null>(
                  ["open-conversation-request"],
                  messageConversationId,
                );
              }
            : undefined,
          currentConversationId: conversationId,
          sharedConversationHeading: `Other chats with ${
            mergedDisplayNameByMri[mri] ||
            part.mentionedDisplayName ||
            part.text.replace(/^@/, "")
          }`,
          sharedConversations: sharedConversationsByMri[mri] ?? [],
        };
      },
      [
        conversationId,
        conversationKind,
        queryClient,
        sharedConversationsByMri,
        threadPresenceByMri,
        mergedAvatarFallbackReady,
        mergedAvatarByMri,
        mergedAvatarFullByMri,
        mergedCompanyNameByMri,
        mergedDepartmentByMri,
        mergedDisplayNameByMri,
        mergedEmailByMri,
        mergedJobTitleByMri,
        mergedLocationByMri,
        mergedTenantNameByMri,
      ],
    );

    const {
      highlightedMessageId,
      matchingMessageIds,
      scrollToMessage: scrollToMessageInViewport,
      submitSearch,
    } = useThreadSearch({
      viewportRef,
      displayMessages,
      searchQuery,
      onSearchResultCountChange,
    });
    const scrollToMessage = useCallback(
      (messageId: string) => {
        scrollToMessageInViewport(messageId);
      },
      [scrollToMessageInViewport],
    );

    const tailMessageId =
      displayMessages.length > 0
        ? (displayMessages[displayMessages.length - 1]?.message.id ?? null)
        : null;

    const { loadingOlder, onScroll, setPendingScrollMessageId } =
      useThreadScroll({
        tenantId,
        conversationId,
        queryClient,
        viewportRef,
        topSentinelRef,
        threadLoading,
        threadHasData,
        tailMessageId,
        loadedMessageCount,
        rawMessages,
        scrollToMessage,
      });

    useEffect(() => {
      if (!perfEnabled) return;
      updatePerfSnapshot(`thread:${conversationId}`, {
        rawMessageCount: rawMessages.length,
        displayMessageCount: displayMessages.length,
        blockCount: messageBlocks.length,
        memberCount: threadMembers.length,
        searchMatchCount: matchingMessageIds.length,
        domNodeCount: countDomNodes(viewportRef.current),
        hasCachedData: threadData ? 1 : 0,
        loadingOlder: loadingOlder ? 1 : 0,
      });
    }, [
      conversationId,
      displayMessages.length,
      loadingOlder,
      matchingMessageIds.length,
      messageBlocks.length,
      rawMessages.length,
      threadData,
      threadMembers.length,
      perfEnabled,
    ]);

    useImperativeHandle(
      ref,
      () => ({
        submitSearch(query: string) {
          submitSearch(query);
        },
      }),
      [submitSearch],
    );

    const mergeThreadData = useCallback(
      (incoming: ThreadQueryData) => {
        queryClient.setQueryData<ThreadQueryData>(
          teamsKeys.thread(tenantId, conversationId),
          (old) => {
            if (!old) return incoming;
            const merged = new Map(
              old.messages.map((message) => [message.id, message]),
            );
            for (const message of incoming.messages) {
              merged.set(message.id, message);
            }
            const messages = sortMessagesByTimestamp([...merged.values()]);
            return {
              messages,
              olderPageUrl: old.olderPageUrl ?? incoming.olderPageUrl,
              moreOlder: old.moreOlder || incoming.moreOlder,
            };
          },
        );
      },
      [conversationId, queryClient, tenantId],
    );

    const openMessageRef = useCallback(
      async (targetConversationId: string, messageId: string) => {
        if (targetConversationId !== conversationId) {
          queryClient.setQueryData<string | null>(
            ["open-conversation-request"],
            targetConversationId,
          );
          return;
        }
        if (rawMessages.some((message) => message.id === messageId)) {
          scrollToMessage(messageId);
          return;
        }
        const res = await teamsThreadService.getAnchoredMessages(
          tenantId,
          targetConversationId,
          messageId,
        );
        if (!res) return;
        setPendingScrollMessageId(messageId);
        mergeThreadData(threadQueryDataFromResponse(res));
      },
      [
        conversationId,
        mergeThreadData,
        queryClient,
        rawMessages,
        scrollToMessage,
        setPendingScrollMessageId,
        tenantId,
      ],
    );

    const handleDeleteMessage = useCallback(
      async (targetConversationId: string, messageId: string) => {
        try {
          await teamsThreadService.deleteMessage(
            tenantId,
            targetConversationId,
            messageId,
          );
          queryClient.setQueryData<ThreadQueryData>(
            teamsKeys.thread(tenantId, conversationId),
            (old) => {
              if (!old) return old;
              return {
                ...old,
                messages: old.messages.map((m) =>
                  m.id === messageId ? teamsThreadService.markDeleted(m) : m,
                ),
              };
            },
          );
        } catch (err) {
          console.error("Failed to delete message:", err);
        }
      },
      [conversationId, queryClient, tenantId],
    );

    useEffect(() => {
      if (!autoFocus) return;
      viewportRef.current?.focus({ preventScroll: true });
    }, [autoFocus]);

    return (
      <section
        ref={viewportRef}
        onScroll={onScroll}
        tabIndex={-1}
        className="relative flex-1 overflow-y-auto overflow-x-hidden bg-background outline-none overscroll-y-contain"
        aria-label="Message thread"
        style={{ overflowAnchor: "none" }}
      >
        {loadingOlder ? (
          <div className="pointer-events-none absolute top-3 left-1/2 z-10 -translate-x-1/2">
            <span className="inline-flex items-center gap-2 rounded-full border border-border bg-background/95 px-4 py-1.5 text-[11px] font-medium text-muted-foreground">
              <span className="size-1.5 animate-pulse rounded-full bg-primary/60" />
              Loading older messages…
            </span>
          </div>
        ) : null}
        <div className="bg-background px-0 pt-4 pb-4">
          {threadQueryError ? (
            <div className="flex flex-col items-center gap-4 py-24">
              <div className="flex size-12 items-center justify-center rounded-2xl bg-destructive/10">
                <span className="text-xl">!</span>
              </div>
              <p className="text-[14px] text-muted-foreground">
                Could not load messages
              </p>
              <button
                type="button"
                onClick={() => void refetchThread()}
                className="rounded-xl bg-primary px-4 py-2 text-[13px] font-medium text-primary-foreground transition-colors hover:bg-primary/90"
              >
                Try again
              </button>
            </div>
          ) : threadLoading && !threadHasData ? (
            <div className="space-y-6 py-10">
              {[0.9, 0.55, 0.75].map((w) => (
                <div key={`skel-${w}`} className="flex gap-3 px-3">
                  <div className="size-9 shrink-0 animate-pulse rounded-xl bg-accent" />
                  <div className="flex-1 space-y-2">
                    <div className="h-3.5 w-24 animate-pulse rounded-lg bg-accent" />
                    <div
                      className="h-10 animate-pulse rounded-xl bg-accent"
                      style={{ width: `${w * 100}%` }}
                    />
                  </div>
                </div>
              ))}
            </div>
          ) : rawMessages.length === 0 ? (
            <div className="flex flex-col items-center gap-3 py-24">
              <div className="flex size-16 items-center justify-center rounded-2xl bg-accent">
                <span className="text-2xl text-muted-foreground/30">💬</span>
              </div>
              <p className="text-[14px] text-muted-foreground/50">
                No messages yet
              </p>
            </div>
          ) : displayMessages.length === 0 ? (
            <p className="py-24 text-center text-[14px] text-muted-foreground/40">
              Only meeting and call activity in this thread.
            </p>
          ) : (
            <ul className="flex flex-col" aria-label="Loaded messages">
              <li
                ref={topSentinelRef}
                className="h-px shrink-0 list-none"
                aria-hidden
              />
              {messageBlocks.map((block) =>
                block.kind === "day" ? (
                  <li key={block.key} className="list-none px-3 py-4">
                    <div className="flex items-center gap-4">
                      <div className="h-px flex-1 bg-border" />
                      <span className="rounded-full bg-accent px-3 py-1 text-[11px] font-semibold tracking-wide text-muted-foreground/60 uppercase">
                        {block.label}
                      </span>
                      <div className="h-px flex-1 bg-border" />
                    </div>
                  </li>
                ) : (
                  (() => {
                    const mri = canonAvatarMri(
                      block.entry.message.fromMri || block.entry.message.from,
                    );
                    const messageConversationId = profileMessageConversationId(
                      conversationKind,
                      conversationId,
                      sharedConversationsByMri[mri] ?? [],
                    );
                    const profile: ProfileData | null = block.entry.self
                      ? null
                      : {
                          mri,
                          displayName:
                            mergedDisplayNameByMri[mri] ||
                            block.entry.displayName,
                          avatarThumbSrc: mergedAvatarByMri[mri],
                          avatarFullSrc:
                            mergedAvatarFullByMri[mri] ??
                            mergedAvatarByMri[mri],
                          avatarFallbackReady: mergedAvatarFallbackReady,
                          email: mergedEmailByMri[mri],
                          jobTitle: mergedJobTitleByMri[mri],
                          department: mergedDepartmentByMri[mri],
                          companyName: mergedCompanyNameByMri[mri],
                          tenantName: mergedTenantNameByMri[mri],
                          location: mergedLocationByMri[mri],
                          presence: threadPresenceByMri[mri],
                          onOpenConversation: (
                            targetConversationId: string,
                          ) => {
                            queryClient.setQueryData<string | null>(
                              ["open-conversation-request"],
                              targetConversationId,
                            );
                          },
                          onMessage: messageConversationId
                            ? () => {
                                queryClient.setQueryData<string | null>(
                                  ["open-conversation-request"],
                                  messageConversationId,
                                );
                              }
                            : undefined,
                          currentConversationId: conversationId,
                          sharedConversationHeading: `Other chats with ${
                            mergedDisplayNameByMri[mri] ||
                            block.entry.displayName
                          }`,
                          sharedConversations:
                            sharedConversationsByMri[mri] ?? [],
                        };
                    return (
                      <MessageRow
                        key={block.key}
                        entry={block.entry}
                        showMeta={block.showMeta}
                        avatarSrc={mergedAvatarByMri[mri]}
                        avatarFallbackReady={mergedAvatarFallbackReady}
                        presence={threadPresenceByMri[mri]}
                        profile={profile}
                        isHighlighted={
                          highlightedMessageId === block.entry.message.id
                        }
                        tenantId={tenantId}
                        onOpenMessageRef={openMessageRef}
                        onDeleteMessage={handleDeleteMessage}
                        getMentionProfile={mentionProfileForPart}
                        onOpenProfile={onOpenProfile}
                      />
                    );
                  })()
                ),
              )}
              <li className="h-px shrink-0 list-none" />
            </ul>
          )}
        </div>
      </section>
    );
  },
);
