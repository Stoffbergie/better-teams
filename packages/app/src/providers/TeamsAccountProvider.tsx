import {
  extractTokens,
  getAvailableAccounts,
  getCachedSession,
} from "@better-teams/app/services/desktop/runtime";
import { teamsKeys } from "@better-teams/app/services/teams/query-keys";
import { teamsSessionService } from "@better-teams/app/services/teams/session";
import type {
  TeamsAccountOption,
  TeamsSessionInfo,
} from "@better-teams/core/teams/types";
import { useQuery, useQueryClient } from "@tanstack/react-query";
import {
  createContext,
  type ReactNode,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
} from "react";

const PREFERRED_TENANT_STORAGE_KEY = "better-teams-preferred-tenant-id";
const CACHED_ACCOUNTS_STORAGE_KEY = "better-teams-cached-accounts";
const CACHED_SESSION_STORAGE_KEY = "better-teams-cached-session";

type TeamsAccountContextValue = {
  accounts: TeamsAccountOption[];
  activeTenantId?: string;
  selectedTenantId?: string | null;
  pendingTenantId?: string | null;
  isSwitchingAccount: boolean;
  activeSession?: TeamsSessionInfo;
  switchAccount: (tenantId: string | null) => void;
  persistedPreference: string | null;
};

const TeamsAccountContext = createContext<TeamsAccountContextValue | null>(
  null,
);

function readPreferredTenantId(): string | null {
  try {
    return localStorage.getItem(PREFERRED_TENANT_STORAGE_KEY);
  } catch {
    return null;
  }
}

function writePreferredTenantId(tenantId: string | null): void {
  try {
    if (tenantId) {
      localStorage.setItem(PREFERRED_TENANT_STORAGE_KEY, tenantId);
      return;
    }
    localStorage.removeItem(PREFERRED_TENANT_STORAGE_KEY);
  } catch {}
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function normalizeAccounts(
  accounts: TeamsAccountOption[] | undefined,
): TeamsAccountOption[] {
  return [...(accounts ?? [])].sort((a, b) =>
    (a.upn ?? "").localeCompare(b.upn ?? ""),
  );
}

async function loadAvailableAccounts(): Promise<TeamsAccountOption[]> {
  const accounts = normalizeAccounts(await getAvailableAccounts());
  if (accounts.length > 0) return accounts;
  const tokens = await extractTokens();
  const byTenant = new Map<string, TeamsAccountOption>();
  for (const token of tokens) {
    const key = token.tenantId ?? token.upn;
    if (!key) continue;
    byTenant.set(key, {
      upn: token.upn,
      tenantId: token.tenantId,
    });
  }
  return normalizeAccounts([...byTenant.values()]);
}

function readCachedAccounts(): TeamsAccountOption[] | undefined {
  try {
    const raw = localStorage.getItem(CACHED_ACCOUNTS_STORAGE_KEY);
    if (!raw) return undefined;
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return undefined;
    return normalizeAccounts(
      parsed.filter(isRecord).flatMap((account) => {
        const upn = typeof account.upn === "string" ? account.upn : undefined;
        const tenantId =
          typeof account.tenantId === "string" ? account.tenantId : undefined;
        return upn || tenantId ? [{ upn, tenantId }] : [];
      }),
    );
  } catch {
    return undefined;
  }
}

function writeCachedAccounts(accounts: TeamsAccountOption[]): void {
  if (accounts.length === 0) return;
  try {
    localStorage.setItem(CACHED_ACCOUNTS_STORAGE_KEY, JSON.stringify(accounts));
  } catch {}
}

function readCachedSession(): TeamsSessionInfo | undefined {
  try {
    const raw = localStorage.getItem(CACHED_SESSION_STORAGE_KEY);
    if (!raw) return undefined;
    const parsed = JSON.parse(raw);
    if (!isRecord(parsed)) return undefined;
    const tenantId = parsed.tenantId;
    if (typeof tenantId !== "string") return undefined;
    return {
      upn: typeof parsed.upn === "string" ? parsed.upn : undefined,
      tenantId,
      skypeId: typeof parsed.skypeId === "string" ? parsed.skypeId : undefined,
      expiresAt:
        typeof parsed.expiresAt === "string" || parsed.expiresAt === null
          ? parsed.expiresAt
          : null,
      region: typeof parsed.region === "string" ? parsed.region : null,
    };
  } catch {
    return undefined;
  }
}

function writeCachedSession(session: TeamsSessionInfo | undefined): void {
  if (!session) return;
  try {
    localStorage.setItem(CACHED_SESSION_STORAGE_KEY, JSON.stringify(session));
  } catch {}
}

function sessionTenantForSelection(
  session: TeamsSessionInfo | undefined,
): string | null {
  if (!session || session.tenantId === "__default__") return null;
  return session.tenantId;
}

function resolveSelectedTenantId(
  accounts: TeamsAccountOption[],
  preferredTenantId: string | null,
): string | undefined {
  if (
    preferredTenantId &&
    accounts.some((account) => account.tenantId === preferredTenantId)
  ) {
    return preferredTenantId;
  }
  return accounts[0]?.tenantId ?? preferredTenantId ?? undefined;
}

async function initializeTeamsSession(
  tenantId?: string | null,
): Promise<TeamsSessionInfo> {
  const cached = await getCachedSession(tenantId);
  if (cached?.tenantId) return cached;
  return teamsSessionService.initialize(tenantId);
}

export function TeamsAccountProvider({ children }: { children: ReactNode }) {
  const queryClient = useQueryClient();

  const [initialAccounts] = useState(() => readCachedAccounts());
  const [cachedSession, setCachedSession] = useState(() => readCachedSession());
  const [persistedPreference, setPersistedPreference] = useState<string | null>(
    () => readPreferredTenantId(),
  );
  const [pendingTenantId, setPendingTenantId] = useState<string | null>(null);

  const { data: accountsData } = useQuery({
    queryKey: teamsKeys.accounts(),
    queryFn: loadAvailableAccounts,
    initialData: initialAccounts,
    initialDataUpdatedAt: initialAccounts ? 0 : undefined,
    staleTime: 30_000,
    gcTime: Number.POSITIVE_INFINITY,
  });

  const accounts = useMemo(
    () => normalizeAccounts(accountsData),
    [accountsData],
  );

  const selectedTenantId = resolveSelectedTenantId(
    accounts,
    persistedPreference ?? sessionTenantForSelection(cachedSession),
  );
  const cachedSessionForSelection =
    sessionTenantForSelection(cachedSession) === (selectedTenantId ?? null)
      ? cachedSession
      : undefined;

  const { data: sessionData } = useQuery({
    queryKey: teamsKeys.session(selectedTenantId),
    queryFn: async () => initializeTeamsSession(selectedTenantId),
    initialData: cachedSessionForSelection,
    initialDataUpdatedAt: cachedSessionForSelection ? 0 : undefined,
    enabled: true,
    staleTime: 30_000,
    gcTime: Number.POSITIVE_INFINITY,
  });

  useEffect(() => {
    writeCachedAccounts(accounts);
  }, [accounts]);

  useEffect(() => {
    if (!sessionData) return;
    writeCachedSession(sessionData);
    setCachedSession(sessionData);
  }, [sessionData]);

  const switchAccount = useCallback(
    (tenantId: string | null) => {
      const nextTenantId = tenantId ?? null;
      if (nextTenantId === selectedTenantId) return;
      const previousTenantId = persistedPreference;
      setPersistedPreference(nextTenantId);
      setPendingTenantId(nextTenantId);
      teamsSessionService.clearTenantClient(nextTenantId);
      void queryClient
        .fetchQuery({
          queryKey: teamsKeys.session(nextTenantId),
          queryFn: () => initializeTeamsSession(nextTenantId),
          staleTime: 30_000,
        })
        .then(() => {
          writePreferredTenantId(nextTenantId);
          setPendingTenantId((current) =>
            current === nextTenantId ? null : current,
          );
        })
        .catch(() => {
          setPersistedPreference(previousTenantId);
          teamsSessionService.clearTenantClient(nextTenantId);
          setPendingTenantId((current) =>
            current === nextTenantId ? null : current,
          );
        });
    },
    [persistedPreference, queryClient, selectedTenantId],
  );

  const activeTenantId =
    selectedTenantId ?? sessionTenantForSelection(sessionData) ?? undefined;

  const value = useMemo<TeamsAccountContextValue>(
    () => ({
      accounts,
      activeTenantId,
      selectedTenantId,
      pendingTenantId,
      isSwitchingAccount:
        pendingTenantId != null && pendingTenantId === selectedTenantId,
      activeSession: sessionData,
      switchAccount,
      persistedPreference,
    }),
    [
      accounts,
      activeTenantId,
      pendingTenantId,
      persistedPreference,
      selectedTenantId,
      sessionData,
      switchAccount,
    ],
  );

  return (
    <TeamsAccountContext.Provider value={value}>
      {children}
    </TeamsAccountContext.Provider>
  );
}

export function useTeamsAccountContext(): TeamsAccountContextValue {
  const value = useContext(TeamsAccountContext);
  if (!value) {
    throw new Error(
      "useTeamsAccountContext must be used within TeamsAccountProvider",
    );
  }
  return value;
}
