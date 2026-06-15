import { preloadConversationThread } from "@better-teams/app/features/chat/thread/preload";
import {
  useActiveTeamsAccount,
  useMaintainTeamsAvailability,
  useTeamsConversations,
  useTeamsPresence,
  useTeamsProfilePresentation,
  useTeamsSession,
} from "@better-teams/app/features/chat/workspace/teams-hooks";
import {
  beginPerfMeasure,
  countDomNodes,
  isPerfEnabled,
  recordPerfMetric,
  updatePerfSnapshot,
} from "@better-teams/app/platform/perf";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import { teamsThreadService } from "@better-teams/app/services/teams/thread";
import { canonAvatarMri } from "@better-teams/core/teams/profile/avatars";
import type {
  Conversation,
  ConversationMember,
} from "@better-teams/core/teams/types";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import {
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
  useTransition,
} from "react";
import type { ComposerMentionCandidate } from "../composer/Composer";
import type { ProfileData } from "../profile/ProfileCard";
import type { ThreadViewHandle } from "../thread/ThreadView";
import {
  useConversationHoverPrefetch,
  useFavoriteConversationMutation,
  useSharedConversationLookup,
  useSidebarConversationViewModel,
} from "./workspace-hooks";

type SelectionFocusTarget = "sidebar" | "thread" | "composer";

function displayNameFromUpn(upn?: string): string | undefined {
  const local = upn?.split("@")[0]?.trim();
  if (!local) return undefined;
  const parts = local
    .split(/[._-]+/)
    .map((part) => part.trim())
    .filter(Boolean);
  if (parts.length === 0) return undefined;
  return parts
    .map((part) => part[0]?.toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");
}

function selfMriFromSkypeId(skypeId?: string): string | null {
  if (!skypeId) return null;
  return canonAvatarMri(skypeId.startsWith("8:") ? skypeId : `8:${skypeId}`);
}

export function useProductivityWorkspaceController() {
  const queryClient = useQueryClient();
  const { activeTenantId, accounts, isSwitchingAccount, switchAccount } =
    useActiveTeamsAccount();
  const sessionQuery = useTeamsSession();
  const session = sessionQuery.session;
  const liveSessionReady = Boolean(session);
  useMaintainTeamsAvailability(liveSessionReady);
  const conversationsQuery = useTeamsConversations(liveSessionReady);
  const conversations = conversationsQuery.conversations;
  const profilePresentation = useTeamsProfilePresentation({
    conversations,
    selfSkypeId: session?.skypeId,
  });
  const presenceByMri = useTeamsPresence({
    conversations,
    selfSkypeId: session?.skypeId,
  });
  const avatarThumbByMri = profilePresentation.avatarThumbs;
  const avatarFullByMri = profilePresentation.avatarFull;
  const displayNameByMri = profilePresentation.displayNames;
  const emailByMri = profilePresentation.emails;
  const jobTitleByMri = profilePresentation.jobTitles;
  const departmentByMri = profilePresentation.departments;
  const companyNameByMri = profilePresentation.companyNames;
  const tenantNameByMri = profilePresentation.tenantNames;
  const locationByMri = profilePresentation.locations;
  const avatarFallbackReady = profilePresentation.avatarFallbackReady;
  const selfMri = useMemo(
    () => selfMriFromSkypeId(session?.skypeId),
    [session?.skypeId],
  );
  const selfDisplayName = useMemo(() => {
    const profileName = selfMri ? displayNameByMri[selfMri]?.trim() : "";
    return profileName || displayNameFromUpn(session?.upn);
  }, [displayNameByMri, selfMri, session?.upn]);

  const [selectedId, setSelectedId] = useState<string | null>(null);
  const [pendingSelectedId, setPendingSelectedId] = useState<string | null>(
    null,
  );
  const [announcement, setAnnouncement] = useState("");
  const [profileSidebarProfile, setProfileSidebarProfile] =
    useState<ProfileData | null>(null);
  const [membersSidebarOpen, setMembersSidebarOpen] = useState(false);
  const [threadSearchQuery, setThreadSearchQuery] = useState("");
  const [threadSearchResultCount, setThreadSearchResultCount] = useState(0);
  const [selectionFocusTarget, setSelectionFocusTarget] =
    useState<SelectionFocusTarget>("thread");
  const perfEnabled = isPerfEnabled();

  const [, startTransition] = useTransition();

  const searchInputRef = useRef<HTMLInputElement>(null);
  const composerRef = useRef<HTMLDivElement>(null);
  const threadViewRef = useRef<ThreadViewHandle>(null);
  const workspaceRef = useRef<HTMLDivElement>(null);
  const pendingSelectionMeasureRef = useRef<ReturnType<
    typeof beginPerfMeasure
  > | null>(null);

  const { allSidebarItems, sidebarItemById, sidebarDisplayNameByMri } =
    useSidebarConversationViewModel({
      conversations,
      selfSkypeId: session?.skypeId,
      avatarThumbByMri,
      avatarFullByMri,
      displayNameByMri,
    });
  const directMessageItemByMri = useMemo(
    () =>
      Object.fromEntries(
        allSidebarItems.flatMap((item) =>
          item.kind === "dm" && item.avatarMri ? [[item.avatarMri, item]] : [],
        ),
      ),
    [allSidebarItems],
  );
  const favoriteMutation = useFavoriteConversationMutation(
    queryClient,
    activeTenantId,
  );
  const profileSidebarMri = profileSidebarProfile?.mri ?? null;
  const { sharedConversationsByMri, loading: sharedConversationsLoading } =
    useSharedConversationLookup({
      activeTenantId,
      profileSidebarMri,
      allSidebarItems,
      sidebarItemById,
      sidebarDisplayNameByMri,
      displayNameByMri,
      emailByMri,
      queryClient,
    });

  const activeConversationId = useMemo(() => {
    if (selectedId == null) return null;
    return conversations.some((conversation) => conversation.id === selectedId)
      ? selectedId
      : null;
  }, [conversations, selectedId]);
  const { data: openConversationRequest } = useQuery({
    queryKey: ["open-conversation-request"],
    queryFn: async () => null as string | null,
    enabled: false,
    initialData: null,
    staleTime: Number.POSITIVE_INFINITY,
    gcTime: Number.POSITIVE_INFINITY,
  });

  const selectedItem = activeConversationId
    ? (sidebarItemById[activeConversationId] ?? null)
    : null;
  const { data: selectedThreadMembers } = useQuery({
    queryKey: activeConversationId
      ? teamsKeys.threadMembers(activeTenantId, activeConversationId)
      : ([
          ...teamsKeys.all,
          "thread-members",
          teamsKeys.scope(activeTenantId),
          "__none__",
        ] as const),
    queryFn: async () => {
      if (!activeConversationId) return [] satisfies ConversationMember[];
      return teamsThreadService.getMembers(
        activeTenantId,
        activeConversationId,
      );
    },
    enabled:
      liveSessionReady &&
      Boolean(activeConversationId) &&
      selectedItem?.kind !== "dm",
    staleTime: 60_000,
  });
  const selectedConversationMembers = useMemo(
    () =>
      selectedItem?.kind !== "dm" && selectedThreadMembers?.length
        ? selectedThreadMembers
        : (selectedItem?.conversation.members ?? []),
    [
      selectedItem?.conversation.members,
      selectedItem?.kind,
      selectedThreadMembers,
    ],
  );
  const selectedProfileConversations = useMemo<Conversation[]>(
    () =>
      selectedItem && selectedConversationMembers.length > 0
        ? [
            {
              id: selectedItem.id,
              members: selectedConversationMembers,
            },
          ]
        : [],
    [selectedConversationMembers, selectedItem],
  );
  const selectedMemberProfilePresentation = useTeamsProfilePresentation({
    conversations: selectedProfileConversations,
    selfSkypeId: session?.skypeId,
  });
  const selectedAvatarFallbackReady =
    avatarFallbackReady &&
    selectedMemberProfilePresentation.avatarFallbackReady;
  const selectedMemberPresenceByMri = useTeamsPresence({
    conversations: selectedProfileConversations,
    selfSkypeId: session?.skypeId,
  });
  const selectedAvatarThumbByMri = useMemo(
    () => ({
      ...avatarThumbByMri,
      ...selectedMemberProfilePresentation.avatarThumbs,
    }),
    [avatarThumbByMri, selectedMemberProfilePresentation.avatarThumbs],
  );
  const selectedAvatarFullByMri = useMemo(
    () => ({
      ...avatarFullByMri,
      ...selectedMemberProfilePresentation.avatarFull,
    }),
    [avatarFullByMri, selectedMemberProfilePresentation.avatarFull],
  );
  const selectedDisplayNameByMri = useMemo(
    () => ({
      ...displayNameByMri,
      ...selectedMemberProfilePresentation.displayNames,
    }),
    [displayNameByMri, selectedMemberProfilePresentation.displayNames],
  );
  const selectedEmailByMri = useMemo(
    () => ({ ...emailByMri, ...selectedMemberProfilePresentation.emails }),
    [emailByMri, selectedMemberProfilePresentation.emails],
  );
  const selectedJobTitleByMri = useMemo(
    () => ({
      ...jobTitleByMri,
      ...selectedMemberProfilePresentation.jobTitles,
    }),
    [jobTitleByMri, selectedMemberProfilePresentation.jobTitles],
  );
  const selectedPresenceByMri = useMemo(
    () => ({ ...presenceByMri, ...selectedMemberPresenceByMri }),
    [presenceByMri, selectedMemberPresenceByMri],
  );
  const selectedMemberProfiles = useMemo<ProfileData[]>(() => {
    const currentConversationId = selectedItem?.id;
    const seen = new Set<string>();

    return selectedConversationMembers.flatMap((member) => {
      const mri = canonAvatarMri(member.id);
      if (!mri || seen.has(mri)) return [];
      seen.add(mri);
      const email = selectedEmailByMri[mri] ?? member.userPrincipalName;
      const displayName =
        selectedDisplayNameByMri[mri]?.trim() ||
        member.displayName?.trim() ||
        member.friendlyName?.trim() ||
        email?.trim() ||
        (mri === selfMri ? selfDisplayName : undefined);
      if (!displayName) return [];

      return [
        {
          mri,
          displayName,
          avatarThumbSrc: selectedAvatarThumbByMri[mri],
          avatarFullSrc:
            selectedAvatarFullByMri[mri] ?? selectedAvatarThumbByMri[mri],
          avatarFallbackReady: selectedAvatarFallbackReady,
          email,
          jobTitle: selectedJobTitleByMri[mri],
          department: departmentByMri[mri],
          companyName: companyNameByMri[mri],
          tenantName: tenantNameByMri[mri],
          location: locationByMri[mri],
          presence: selectedPresenceByMri[mri],
          currentConversationId,
          sharedConversationHeading: `Other chats with ${displayName}`,
          sharedConversations: sharedConversationsByMri[mri] ?? [],
          onOpenConversation: (conversationId: string) => {
            const item = sidebarItemById[conversationId];
            setSelectedId(conversationId);
            setSelectionFocusTarget("thread");
            setAnnouncement(
              item ? `Opened ${item.title}` : "Opened conversation",
            );
            setMembersSidebarOpen(false);
            setProfileSidebarProfile(null);
          },
        },
      ];
    });
  }, [
    companyNameByMri,
    departmentByMri,
    locationByMri,
    selectedAvatarFullByMri,
    selectedAvatarFallbackReady,
    selectedAvatarThumbByMri,
    selectedConversationMembers,
    selectedDisplayNameByMri,
    selectedEmailByMri,
    selectedItem?.id,
    selectedJobTitleByMri,
    selectedPresenceByMri,
    session?.upn,
    selfDisplayName,
    selfMri,
    sharedConversationsByMri,
    sidebarItemById,
    tenantNameByMri,
  ]);
  const selectedHeaderMemberCount = useMemo(() => {
    if (!selectedItem) return null;
    if (selectedItem.kind === "dm") return 2;
    if (selectedMemberProfiles.length > 0) {
      return selectedMemberProfiles.length;
    }
    const summaryCount = selectedItem.conversation.memberCount;
    return typeof summaryCount === "number" && summaryCount > 0
      ? summaryCount
      : null;
  }, [selectedItem, selectedMemberProfiles.length]);
  const selectedHeaderAvatarMris = useMemo(() => {
    const peerMris = selectedMemberProfiles
      .map((profile) => profile.mri)
      .filter((mri) => mri !== selfMri);
    return peerMris.length > 0
      ? peerMris
      : selectedMemberProfiles.map((profile) => profile.mri);
  }, [selectedMemberProfiles, selfMri]);
  const selectedHeaderAvatarLabelsByMri = useMemo(
    () =>
      Object.fromEntries(
        selectedMemberProfiles.map((profile) => [
          profile.mri,
          profile.displayName,
        ]),
      ),
    [selectedMemberProfiles],
  );
  const selectedProfileData = useMemo<ProfileData | null>(() => {
    if (
      !selectedItem ||
      selectedItem.kind !== "dm" ||
      !selectedItem.avatarMri
    ) {
      return null;
    }
    const mri = selectedItem.avatarMri;
    return {
      mri,
      displayName: displayNameByMri[mri] || selectedItem.title,
      avatarThumbSrc: avatarThumbByMri[mri],
      avatarFullSrc: avatarFullByMri[mri] ?? avatarThumbByMri[mri],
      avatarFallbackReady,
      email: emailByMri[mri],
      jobTitle: jobTitleByMri[mri],
      department: departmentByMri[mri],
      companyName: companyNameByMri[mri],
      tenantName: tenantNameByMri[mri],
      location: locationByMri[mri],
      presence: presenceByMri[mri],
      onMessage:
        selectedItem.kind === "dm"
          ? undefined
          : () => {
              setSelectedId(selectedItem.id);
              setSelectionFocusTarget("composer");
              setAnnouncement(
                `Ready to message ${displayNameByMri[mri] || selectedItem.title}`,
              );
              setMembersSidebarOpen(false);
              setProfileSidebarProfile(null);
            },
      onOpenConversation: (conversationId: string) => {
        const item = sidebarItemById[conversationId];
        setSelectedId(conversationId);
        setSelectionFocusTarget("thread");
        setAnnouncement(item ? `Opened ${item.title}` : "Opened conversation");
        setMembersSidebarOpen(false);
        setProfileSidebarProfile(null);
      },
      currentConversationId: selectedItem.id,
      sharedConversationHeading: `Other chats with ${displayNameByMri[mri] || selectedItem.title}`,
      sharedConversations: sharedConversationsByMri[mri] ?? [],
    };
  }, [
    avatarFullByMri,
    avatarFallbackReady,
    avatarThumbByMri,
    companyNameByMri,
    departmentByMri,
    displayNameByMri,
    emailByMri,
    jobTitleByMri,
    locationByMri,
    presenceByMri,
    selectedItem,
    sidebarItemById,
    sharedConversationsByMri,
    tenantNameByMri,
  ]);
  const profileSidebarData = useMemo<ProfileData | null>(() => {
    if (!profileSidebarProfile) return null;
    const mri = profileSidebarProfile.mri;
    const displayName =
      selectedDisplayNameByMri[mri] || profileSidebarProfile.displayName;
    const onOpenConversation = profileSidebarProfile.onOpenConversation
      ? (conversationId: string) => {
          profileSidebarProfile.onOpenConversation?.(conversationId);
          setProfileSidebarProfile(null);
        }
      : undefined;
    const messageItem = directMessageItemByMri[mri];
    const onMessage =
      messageItem && messageItem.id !== activeConversationId
        ? () => {
            setSelectedId(messageItem.id);
            setSelectionFocusTarget("composer");
            setAnnouncement(`Ready to message ${displayName}`);
            setMembersSidebarOpen(false);
            setProfileSidebarProfile(null);
          }
        : profileSidebarProfile.onMessage
          ? () => {
              profileSidebarProfile.onMessage?.();
              setProfileSidebarProfile(null);
            }
          : undefined;
    const sidebarAvatarFallbackReady =
      Boolean(
        selectedAvatarFullByMri[mri] ??
          selectedAvatarThumbByMri[mri] ??
          profileSidebarProfile.avatarFullSrc ??
          profileSidebarProfile.avatarThumbSrc,
      ) || selectedAvatarFallbackReady;

    return {
      ...profileSidebarProfile,
      displayName,
      avatarThumbSrc:
        selectedAvatarThumbByMri[mri] ?? profileSidebarProfile.avatarThumbSrc,
      avatarFullSrc:
        selectedAvatarFullByMri[mri] ??
        selectedAvatarThumbByMri[mri] ??
        profileSidebarProfile.avatarFullSrc,
      avatarFallbackReady: sidebarAvatarFallbackReady,
      email: selectedEmailByMri[mri] ?? profileSidebarProfile.email,
      jobTitle: selectedJobTitleByMri[mri] ?? profileSidebarProfile.jobTitle,
      department: departmentByMri[mri] ?? profileSidebarProfile.department,
      companyName: companyNameByMri[mri] ?? profileSidebarProfile.companyName,
      tenantName: tenantNameByMri[mri] ?? profileSidebarProfile.tenantName,
      location: locationByMri[mri] ?? profileSidebarProfile.location,
      presence: selectedPresenceByMri[mri] ?? profileSidebarProfile.presence,
      onOpenConversation,
      onMessage,
      sharedConversationHeading: `Other chats with ${displayName}`,
      sharedConversationsLoading,
      sharedConversations:
        sharedConversationsByMri[mri] ??
        profileSidebarProfile.sharedConversations ??
        [],
    };
  }, [
    activeConversationId,
    companyNameByMri,
    departmentByMri,
    directMessageItemByMri,
    locationByMri,
    profileSidebarProfile,
    selectedAvatarFallbackReady,
    selectedAvatarFullByMri,
    selectedAvatarThumbByMri,
    selectedDisplayNameByMri,
    selectedEmailByMri,
    selectedJobTitleByMri,
    selectedPresenceByMri,
    sharedConversationsLoading,
    sharedConversationsByMri,
    tenantNameByMri,
  ]);
  const composerMentionCandidates = useMemo<ComposerMentionCandidate[]>(() => {
    const seen = new Set<string>();
    const pushCandidate = (
      mriValue: string | undefined,
      displayName: string | undefined,
      email?: string,
    ) => {
      if (!mriValue || !displayName) return [] as ComposerMentionCandidate[];
      const mri = canonAvatarMri(mriValue);
      if (seen.has(mri)) return [] as ComposerMentionCandidate[];
      seen.add(mri);
      return [
        {
          mri,
          displayName,
          email: selectedEmailByMri[mri] || email,
          avatarSrc: selectedAvatarThumbByMri[mri],
        },
      ];
    };

    const selectedMembers = selectedConversationMembers.flatMap((member) =>
      pushCandidate(
        member.id,
        selectedDisplayNameByMri[canonAvatarMri(member.id)] ||
          member.displayName ||
          member.friendlyName ||
          member.userPrincipalName,
        member.userPrincipalName,
      ),
    );

    const selectedDmPeer =
      selectedItem?.avatarMri && selectedItem.kind === "dm"
        ? pushCandidate(
            selectedItem.avatarMri,
            selectedDisplayNameByMri[selectedItem.avatarMri] ||
              selectedItem.title,
            selectedEmailByMri[selectedItem.avatarMri],
          )
        : [];

    const sidebarPeers = allSidebarItems.flatMap((item) =>
      item.avatarMri
        ? pushCandidate(
            item.avatarMri,
            selectedDisplayNameByMri[item.avatarMri] || item.title,
            selectedEmailByMri[item.avatarMri],
          )
        : [],
    );

    const profileDirectory = Object.entries(selectedDisplayNameByMri).flatMap(
      ([mri, displayName]) =>
        pushCandidate(mri, displayName, selectedEmailByMri[mri]),
    );

    return [
      ...selectedMembers,
      ...selectedDmPeer,
      ...sidebarPeers,
      ...profileDirectory,
    ];
  }, [
    allSidebarItems,
    selectedAvatarThumbByMri,
    selectedDisplayNameByMri,
    selectedEmailByMri,
    selectedItem,
    selectedConversationMembers,
  ]);

  const selfAvatarSrc = selfMri ? avatarThumbByMri[selfMri] : undefined;

  const accountAvatarByTenant = useMemo(
    () =>
      activeTenantId && selfAvatarSrc
        ? { [activeTenantId]: selfAvatarSrc }
        : {},
    [activeTenantId, selfAvatarSrc],
  );
  useEffect(() => {
    const requestedConversationId = openConversationRequest;
    if (!requestedConversationId) return;
    const requestedItem = sidebarItemById[requestedConversationId];
    if (!requestedItem) return;
    setSelectedId(requestedConversationId);
    setSelectionFocusTarget("thread");
    setMembersSidebarOpen(false);
    setProfileSidebarProfile(null);
    queryClient.setQueryData(["open-conversation-request"], null);
  }, [openConversationRequest, queryClient, sidebarItemById]);

  useEffect(() => {
    if (selectionFocusTarget !== "composer") return;
    composerRef.current?.focus();
  }, [selectionFocusTarget]);

  useEffect(() => {
    if (!perfEnabled) return;
    updatePerfSnapshot("workspace.sidebar", {
      conversationCount: allSidebarItems.length,
      favoriteCount: allSidebarItems.filter((item) => item.isFavorite).length,
      dmCount: allSidebarItems.filter((item) => item.kind === "dm").length,
      groupCount: allSidebarItems.filter((item) => item.kind === "group")
        .length,
      meetingCount: allSidebarItems.filter((item) => item.kind === "meeting")
        .length,
      domNodeCount: countDomNodes(workspaceRef.current),
      selectedConversation:
        activeConversationId ?? pendingSelectedId ?? "__none__",
    });
  }, [activeConversationId, allSidebarItems, pendingSelectedId, perfEnabled]);

  useEffect(() => {
    if (!activeConversationId || !pendingSelectionMeasureRef.current) return;
    pendingSelectionMeasureRef.current({
      conversationId: activeConversationId,
      focusTarget: selectionFocusTarget,
    });
    pendingSelectionMeasureRef.current = null;
  }, [activeConversationId, selectionFocusTarget]);

  const handleSelectConversation = useCallback(
    (id: string, focus: SelectionFocusTarget) => {
      const item = sidebarItemById[id];
      if (liveSessionReady) {
        void queryClient.prefetchQuery({
          queryKey: teamsKeys.thread(activeTenantId, id),
          queryFn: () =>
            preloadConversationThread(activeTenantId ?? undefined, id, 60_000),
          staleTime: 25_000,
        });
      }
      recordPerfMetric("workspace.selectConversation.requested", {
        conversationId: id,
        focusTarget: focus,
      });
      pendingSelectionMeasureRef.current = beginPerfMeasure(
        "workspace.selectConversation",
        {
          conversationId: id,
          focusTarget: focus,
        },
      );
      setPendingSelectedId(id);
      setSelectionFocusTarget(focus);
      startTransition(() => {
        setSelectedId(id);
        setPendingSelectedId(null);
        setThreadSearchQuery("");
        setThreadSearchResultCount(0);
        setAnnouncement(item ? `Opened ${item.title}` : "Opened conversation");
        setMembersSidebarOpen(false);
        setProfileSidebarProfile(null);
      });
    },
    [activeTenantId, liveSessionReady, queryClient, sidebarItemById],
  );

  const { handleHoverConversation, handleHoverConversationEnd } =
    useConversationHoverPrefetch({
      activeTenantId,
      liveSessionReady,
      activeConversationId,
      queryClient,
    });

  const { mutate: mutateFavorite } = favoriteMutation;
  const handleToggleFavorite = useCallback(
    (conversationId: string, favorite: boolean) => {
      mutateFavorite({ conversationId, favorite });
    },
    [mutateFavorite],
  );

  const handleSubmitSearch = useCallback((query: string) => {
    const trimmedQuery = query.trim();
    setThreadSearchQuery(trimmedQuery);
    if (!trimmedQuery) return;
    threadViewRef.current?.submitSearch(trimmedQuery);
  }, []);
  const handleCloseSearch = useCallback(() => {
    setThreadSearchQuery("");
    setThreadSearchResultCount(0);
  }, []);
  const handleOpenSelectedProfile = selectedProfileData
    ? () => {
        setMembersSidebarOpen(false);
        setProfileSidebarProfile(selectedProfileData);
      }
    : undefined;
  const handleOpenThreadProfile = useCallback((profile: ProfileData) => {
    setMembersSidebarOpen(false);
    setProfileSidebarProfile(profile);
  }, []);
  const handleOpenMembersSidebar = useCallback(() => {
    setProfileSidebarProfile(null);
    setMembersSidebarOpen(true);
  }, []);
  const handleOpenMemberProfile = useCallback((profile: ProfileData) => {
    setMembersSidebarOpen(false);
    setProfileSidebarProfile(profile);
  }, []);
  const selectedProfileButtonLabel = selectedProfileData
    ? `Open profile for ${selectedProfileData.displayName}`
    : undefined;

  const accountLoading = !session && sessionQuery.isPending;
  const conversationsLoading =
    allSidebarItems.length === 0 &&
    (!session || conversationsQuery.isPending || conversationsQuery.isFetching);
  const errorMessage =
    !session && sessionQuery.isError
      ? sessionQuery.error instanceof Error
        ? sessionQuery.error.message
        : "Could not connect to Teams"
      : null;

  return {
    activeTenantId,
    accounts,
    accountAvatarByTenant,
    accountLoading,
    activeConversationId,
    allSidebarItems,
    announcement,
    avatarFallbackReady,
    avatarFullByMri,
    avatarThumbByMri,
    companyNameByMri,
    composerMentionCandidates,
    composerRef,
    conversationsLoading,
    departmentByMri,
    displayNameByMri,
    emailByMri,
    errorMessage,
    handleCloseProfileSidebar: () => setProfileSidebarProfile(null),
    handleCloseMembersSidebar: () => setMembersSidebarOpen(false),
    handleCloseSearch,
    handleHoverConversation,
    handleHoverConversationEnd,
    handleOpenMemberProfile,
    handleOpenMembersSidebar,
    handleOpenThreadProfile,
    handleOpenSelectedProfile,
    handleSelectConversation,
    handleSubmitSearch,
    handleToggleFavorite,
    isSwitchingAccount,
    jobTitleByMri,
    liveSessionReady,
    locationByMri,
    membersSidebarOpen,
    pendingSelectedId,
    presenceByMri,
    profileSidebarData,
    searchInputRef,
    selectedAvatarThumbByMri,
    selectedAvatarFallbackReady,
    selectedHeaderAvatarLabelsByMri,
    selectedHeaderAvatarMris,
    selectedItem,
    selectedConversationMembers,
    selectedHeaderMemberCount,
    selectedMemberProfiles,
    selectedPresenceByMri,
    selectedProfileButtonLabel,
    selectionFocusTarget,
    selfDisplayName,
    selfAvatarSrc,
    session,
    sessionQuery,
    setThreadSearchQuery,
    setThreadSearchResultCount,
    sharedConversationsByMri,
    switchAccount,
    tenantNameByMri,
    threadSearchQuery,
    threadSearchResultCount,
    threadViewRef,
    workspaceRef,
  };
}
