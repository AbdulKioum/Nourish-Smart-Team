import { createClient } from "@supabase/supabase-js";
import { get, set } from "idb-keyval";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

export const rawSupabase = createClient(supabaseUrl, supabaseAnonKey, {
  auth: {
    persistSession: true,
    autoRefreshToken: true,
    detectSessionInUrl: true,
    storage: window.localStorage,
  },
});

export const SYNC_TABLES = [
  "users",
  "companies",
  "agents",
  "farmers",
  "doc_agents",
  "doc_supplies",
  "doc_trackings",
  "vet_visits",
  "vaccines",
];

export interface SyncMutation {
  id: string;
  table: string;
  action: "insert" | "update" | "delete";
  payload?: any;
  filters?: { type: "eq" | "in"; col: string; val: any }[];
  timestamp: number;
}

const QUEUE_KEY = "fams_offline_mutation_queue";

// 1. Pull all tables from Supabase to Local IndexedDB
export async function pullAllTablesToLocal(): Promise<boolean> {
  if (!navigator.onLine) return false;
  try {
    await Promise.all(
      SYNC_TABLES.map(async (table) => {
        const { data, error } = await rawSupabase.from(table).select("*");
        if (!error && data) {
          await set(`table_${table}`, data);
        }
      })
    );
    await set("fams_last_synced_at", new Date().toISOString());
    window.dispatchEvent(new CustomEvent("fams-data-synced"));
    return true;
  } catch (err) {
    console.error("Failed to pull tables:", err);
    return false;
  }
}

// 2. Push pending offline inputs to Supabase when online
let isSyncingQueue = false;
export async function processMutationQueue(): Promise<void> {
  if (!navigator.onLine || isSyncingQueue) return;
  isSyncingQueue = true;

  try {
    const queue: SyncMutation[] = (await get(QUEUE_KEY)) || [];
    if (queue.length === 0) {
      isSyncingQueue = false;
      return;
    }

    const remainingQueue: SyncMutation[] = [];

    for (const item of queue) {
      try {
        let query: any = rawSupabase.from(item.table);

        if (item.action === "insert") {
          const rows = Array.isArray(item.payload)
            ? item.payload
            : [item.payload];
          const cleanRows = rows.map((r: any) => {
            const copy = { ...r };
            if (typeof copy.id === "string" && copy.id.startsWith("local_")) {
              delete copy.id;
            }
            return copy;
          });
          query = query.insert(cleanRows);
        } else if (item.action === "update") {
          const copy = { ...item.payload };
          delete copy.id;
          query = query.update(copy);
          for (const f of item.filters || []) {
            if (f.type === "eq") query = query.eq(f.col, f.val);
            if (f.type === "in") query = query.in(f.col, f.val);
          }
        } else if (item.action === "delete") {
          query = query.delete();
          for (const f of item.filters || []) {
            if (f.type === "eq") query = query.eq(f.col, f.val);
            if (f.type === "in") query = query.in(f.col, f.val);
          }
        }

        const { error } = await query;
        if (error) {
          console.error(`Sync error on ${item.table}:`, error);
          // Network error hole abar queue te rakhbo
          if (
            error.message?.includes("fetch") ||
            error.message?.includes("network")
          ) {
            remainingQueue.push(item);
          }
        }
      } catch (err) {
        remainingQueue.push(item);
      }
    }

    await set(QUEUE_KEY, remainingQueue);
    window.dispatchEvent(new CustomEvent("fams-queue-updated"));

    if (remainingQueue.length < queue.length) {
      await pullAllTablesToLocal();
    }
  } finally {
    isSyncingQueue = false;
  }
}

async function addToMutationQueue(
  mutation: Omit<SyncMutation, "id" | "timestamp">
) {
  const queue: SyncMutation[] = (await get(QUEUE_KEY)) || [];
  queue.push({
    ...mutation,
    id: `mut_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`,
    timestamp: Date.now(),
  });
  await set(QUEUE_KEY, queue);
  window.dispatchEvent(new CustomEvent("fams-queue-updated"));
}

// 3. Offline-First Query Builder (Intercepts .from().select()/.insert()/.update()/.delete())
class OfflineQueryBuilder {
  private table: string;
  private action: "select" | "insert" | "update" | "delete" = "select";
  private selectQuery = "*";
  private payload: any = null;
  private filters: { type: "eq" | "in" | "gte" | "lte"; col: string; val: any }[] = [];
  private orderBy: { col: string; ascending: boolean } | null = null;
  private isSingle = false;

  constructor(table: string) {
    this.table = table;
  }

  select(columns = "*") {
    if (this.action === "select") {
      this.selectQuery = columns;
    }
    return this;
  }

  insert(values: any) {
    this.action = "insert";
    this.payload = values;
    return this;
  }

  update(values: any) {
    this.action = "update";
    this.payload = values;
    return this;
  }

  delete() {
    this.action = "delete";
    return this;
  }

  eq(col: string, val: any) {
    this.filters.push({ type: "eq", col, val });
    return this;
  }

  in(col: string, val: any[]) {
    this.filters.push({ type: "in", col, val });
    return this;
  }

  gte(col: string, val: any) {
    this.filters.push({ type: "gte", col, val });
    return this;
  }

  lte(col: string, val: any) {
    this.filters.push({ type: "lte", col, val });
    return this;
  }

  order(col: string, opts?: { ascending?: boolean }) {
    this.orderBy = { col, ascending: opts?.ascending ?? true };
    return this;
  }

  single() {
    this.isSingle = true;
    return this;
  }

  then<TResult1 = any, TResult2 = never>(
    onfulfilled?: ((value: any) => TResult1 | PromiseLike<TResult1>) | null,
    onrejected?: ((reason: any) => TResult2 | PromiseLike<TResult2>) | null
  ): Promise<TResult1 | TResult2> {
    return this.execute().then(onfulfilled, onrejected);
  }

  private matchesFilters(row: any): boolean {
    for (const f of this.filters) {
      const rowVal = row[f.col];
      if (f.type === "eq" && String(rowVal) !== String(f.val)) return false;
      if (f.type === "in") {
        const strArr = (f.val || []).map(String);
        if (!strArr.includes(String(rowVal))) return false;
      }
      if (f.type === "gte" && rowVal < f.val) return false;
      if (f.type === "lte" && rowVal > f.val) return false;
    }
    return true;
  }

  private async attachRelations(rows: any[]): Promise<any[]> {
    if (this.selectQuery.includes("companies(")) {
      const companies: any[] = (await get("table_companies")) || [];
      const compMap = new Map(companies.map((c) => [String(c.id), c]));
      return rows.map((r) => ({
        ...r,
        companies: compMap.get(String(r.company_id)) || null,
      }));
    }
    return rows;
  }

  private async execute(): Promise<{ data: any; error: any }> {
    const storageKey = `table_${this.table}`;

    // =====================================
    // READ (SELECT)
    // =====================================
    if (this.action === "select") {
      if (navigator.onLine) {
        try {
          let q: any = rawSupabase.from(this.table).select(this.selectQuery);
          for (const f of this.filters) {
            if (f.type === "eq") q = q.eq(f.col, f.val);
            if (f.type === "in") q = q.in(f.col, f.val);
            if (f.type === "gte") q = q.gte(f.col, f.val);
            if (f.type === "lte") q = q.lte(f.col, f.val);
          }
          if (this.orderBy) {
            q = q.order(this.orderBy.col, {
              ascending: this.orderBy.ascending,
            });
          }
          if (this.isSingle) {
            q = q.single();
          }

          const { data, error } = await q;
          if (!error && data !== null) {
            if (
              this.filters.length === 0 &&
              !this.isSingle &&
              Array.isArray(data)
            ) {
              await set(storageKey, data);
            }
            return { data, error: null };
          }
        } catch {
          // Net drops mid-request -> Fall back to Local IndexedDB
        }
      }

      // OFFLINE READ FROM INDEXEDDB
      let localRows: any[] = (await get(storageKey)) || [];
      localRows = localRows.filter((r) => this.matchesFilters(r));
      localRows = await this.attachRelations(localRows);

      if (this.orderBy) {
        const { col, ascending } = this.orderBy;
        localRows.sort((a, b) => {
          const valA = a[col] ?? "";
          const valB = b[col] ?? "";
          if (valA < valB) return ascending ? -1 : 1;
          if (valA > valB) return ascending ? 1 : -1;
          return 0;
        });
      }

      if (this.isSingle) {
        return { data: localRows[0] || null, error: null };
      }
      return { data: localRows, error: null };
    }

    // =====================================
    // WRITE (INSERT / UPDATE / DELETE)
    // =====================================
    if (navigator.onLine) {
      try {
        let q: any = rawSupabase.from(this.table);
        if (this.action === "insert") q = q.insert(this.payload).select();
        if (this.action === "update") {
          q = q.update(this.payload);
          for (const f of this.filters) {
            if (f.type === "eq") q = q.eq(f.col, f.val);
            if (f.type === "in") q = q.in(f.col, f.val);
          }
          q = q.select();
        }
        if (this.action === "delete") {
          q = q.delete();
          for (const f of this.filters) {
            if (f.type === "eq") q = q.eq(f.col, f.val);
            if (f.type === "in") q = q.in(f.col, f.val);
          }
        }
        if (this.isSingle && this.action !== "delete") q = q.single();

        const { data, error } = await q;
        if (!error) {
          await this.applyLocalMutation(data);
          return { data, error: null };
        }
      } catch {
        // Failed online write -> Proceed to offline local write & queue
      }
    }

    // OFFLINE WRITE TO LOCAL STORAGE + QUEUE
    const mutatedData = await this.applyLocalMutation();
    await addToMutationQueue({
      table: this.table,
      action: this.action,
      payload: this.payload,
      filters: this.filters.filter(
        (f): f is { type: "eq" | "in"; col: string; val: any } =>
          f.type === "eq" || f.type === "in"
      ),
    });

    return {
      data: this.isSingle ? mutatedData[0] || null : mutatedData,
      error: null,
    };
  }

  private async applyLocalMutation(serverData?: any): Promise<any[]> {
    const storageKey = `table_${this.table}`;
    let localRows: any[] = (await get(storageKey)) || [];
    let affected: any[] = [];

    if (this.action === "insert") {
      if (serverData) {
        const sRows = Array.isArray(serverData) ? serverData : [serverData];
        localRows = [...sRows, ...localRows];
        affected = sRows;
      } else {
        const rawItems = Array.isArray(this.payload)
          ? this.payload
          : [this.payload];
        const newItems = rawItems.map((item) => ({
          ...item,
          id:
            item.id ??
            (this.table === "farmers" || this.table === "agents"
              ? Math.floor(Date.now() / 1000) + Math.floor(Math.random() * 1000)
              : `local_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`),
          created_at: item.created_at || new Date().toISOString(),
        }));
        localRows = [...newItems, ...localRows];
        affected = newItems;
      }
    } else if (this.action === "update") {
      localRows = localRows.map((row) => {
        if (this.matchesFilters(row)) {
          const updated = { ...row, ...this.payload };
          affected.push(updated);
          return updated;
        }
        return row;
      });
    } else if (this.action === "delete") {
      localRows = localRows.filter((row) => !this.matchesFilters(row));
    }

    await set(storageKey, localRows);
    return affected;
  }
}

// Offline-aware Auth Wrapper
const offlineAuth = new Proxy(rawSupabase.auth, {
  get(target, prop, receiver) {
    if (prop === "getUser") {
      return async () => {
        if (navigator.onLine) {
          try {
            const res = await target.getUser();
            if (res.data?.user) {
              await set("fams_cached_user", res.data.user);
              return res;
            }
          } catch {
            // Fall back to cached user
          }
        }
        const cachedUser = await get("fams_cached_user");
        return { data: { user: cachedUser || null }, error: null };
      };
    }
    if (prop === "getSession") {
      return async () => {
        if (navigator.onLine) {
          try {
            const res = await target.getSession();
            if (res.data?.session) {
              await set("fams_cached_session", res.data.session);
              await set("fams_cached_user", res.data.session.user);
              return res;
            }
          } catch {
            // Fall back to cached session
          }
        }
        const cachedSession = await get("fams_cached_session");
        return { data: { session: cachedSession || null }, error: null };
      };
    }
    const val = Reflect.get(target, prop, receiver);
    return typeof val === "function" ? val.bind(target) : val;
  },
});

export const supabase = new Proxy(rawSupabase, {
  get(target, prop, receiver) {
    if (prop === "from") {
      return (table: string) => new OfflineQueryBuilder(table);
    }
    if (prop === "auth") {
      return offlineAuth;
    }
    const value = Reflect.get(target, prop, receiver);
    return typeof value === "function" ? value.bind(target) : value;
  },
});