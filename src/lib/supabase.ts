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
  "sub_agents",
  "retailers",
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
  action: "insert" | "update" | "delete" | "upsert";
  payload?: any;
  filters?: { type: string; col: string; val: any }[];
  timestamp: number;
}

const QUEUE_KEY = "fams_offline_mutation_queue";

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

        if (item.action === "insert" || item.action === "upsert") {
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
          query =
            item.action === "upsert"
              ? query.upsert(cleanRows)
              : query.insert(cleanRows);
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
          if (
            error.message?.includes("fetch") ||
            error.message?.includes("network")
          ) {
            remainingQueue.push(item);
          }
        }
      } catch {
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

class OfflineQueryBuilder {
  private table: string;
  private action: "select" | "insert" | "update" | "delete" | "upsert" = "select";
  private selectQuery = "*";
  private selectOpts: any = undefined;
  private payload: any = null;
  private upsertOpts: any = undefined;
  private filters: { type: string; col: string; val: any }[] = [];
  private orderBys: { col: string; ascending: boolean }[] = [];
  private limitCount: number | null = null;
  private rangeBounds: { from: number; to: number } | null = null;
  private isSingle = false;
  private isMaybeSingle = false;

  constructor(table: string) {
    this.table = table;
  }

  select(columns = "*", opts?: any) {
    if (this.action === "select") {
      this.selectQuery = columns;
      this.selectOpts = opts;
    }
    return this;
  }

  insert(values: any) {
    this.action = "insert";
    this.payload = values;
    return this;
  }

  upsert(values: any, opts?: any) {
    this.action = "upsert";
    this.payload = values;
    this.upsertOpts = opts;
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

  neq(col: string, val: any) {
    this.filters.push({ type: "neq", col, val });
    return this;
  }

  in(col: string, val: any[]) {
    this.filters.push({ type: "in", col, val });
    return this;
  }

  ilike(col: string, val: string) {
    this.filters.push({ type: "ilike", col, val });
    return this;
  }

  like(col: string, val: string) {
    this.filters.push({ type: "ilike", col, val });
    return this;
  }

  is(col: string, val: any) {
    this.filters.push({ type: "is", col, val });
    return this;
  }

  gt(col: string, val: any) {
    this.filters.push({ type: "gt", col, val });
    return this;
  }

  gte(col: string, val: any) {
    this.filters.push({ type: "gte", col, val });
    return this;
  }

  lt(col: string, val: any) {
    this.filters.push({ type: "lt", col, val });
    return this;
  }

  lte(col: string, val: any) {
    this.filters.push({ type: "lte", col, val });
    return this;
  }

  or(filterStr: string) {
    this.filters.push({ type: "or", col: "", val: filterStr });
    return this;
  }

  match(queryObj: Record<string, any>) {
    Object.entries(queryObj).forEach(([k, v]) => {
      this.filters.push({ type: "eq", col: k, val: v });
    });
    return this;
  }

  order(col: string, opts?: { ascending?: boolean }) {
    this.orderBys.push({ col, ascending: opts?.ascending ?? true });
    return this;
  }

  limit(count: number) {
    this.limitCount = count;
    return this;
  }

  range(from: number, to: number) {
    this.rangeBounds = { from, to };
    return this;
  }

  single() {
    this.isSingle = true;
    return this;
  }

  maybeSingle() {
    this.isMaybeSingle = true;
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
      if (f.type === "eq" && String(rowVal ?? "") !== String(f.val ?? ""))
        return false;
      if (f.type === "neq" && String(rowVal ?? "") === String(f.val ?? ""))
        return false;
      if (f.type === "is" && rowVal !== f.val) return false;
      if (f.type === "in") {
        const strArr = (f.val || []).map(String);
        if (!strArr.includes(String(rowVal))) return false;
      }
      if (f.type === "ilike") {
        const cleanPattern = String(f.val || "")
          .replace(/%/g, "")
          .toLowerCase();
        if (
          !String(rowVal || "")
            .toLowerCase()
            .includes(cleanPattern)
        )
          return false;
      }
      if (f.type === "gt" && !(rowVal > f.val)) return false;
      if (f.type === "gte" && !(rowVal >= f.val)) return false;
      if (f.type === "lt" && !(rowVal < f.val)) return false;
      if (f.type === "lte" && !(rowVal <= f.val)) return false;
    }
    return true;
  }

  private async attachRelations(rows: any[]): Promise<any[]> {
    let result = [...rows];
    if (this.selectQuery.includes("companies(")) {
      const companies: any[] = (await get("table_companies")) || [];
      const compMap = new Map(companies.map((c) => [String(c.id), c]));
      result = result.map((r) => ({
        ...r,
        companies: compMap.get(String(r.company_id)) || null,
      }));
    }
    if (this.selectQuery.includes("agents(")) {
      const agents: any[] = (await get("table_agents")) || [];
      const agMap = new Map(agents.map((a) => [String(a.id), a]));
      result = result.map((r) => ({
        ...r,
        agents: agMap.get(String(r.agent_id)) || null,
      }));
    }
    if (this.selectQuery.includes("farmers(")) {
      const farmers: any[] = (await get("table_farmers")) || [];
      const frmMap = new Map(farmers.map((f) => [String(f.id), f]));
      result = result.map((r) => ({
        ...r,
        farmers: frmMap.get(String(r.farmer_id)) || null,
      }));
    }
    return result;
  }

  private applyFiltersToSupabaseQuery(q: any) {
    let query = q;
    for (const f of this.filters) {
      if (f.type === "eq") query = query.eq(f.col, f.val);
      else if (f.type === "neq") query = query.neq(f.col, f.val);
      else if (f.type === "in") query = query.in(f.col, f.val);
      else if (f.type === "ilike") query = query.ilike(f.col, f.val);
      else if (f.type === "is") query = query.is(f.col, f.val);
      else if (f.type === "gt") query = query.gt(f.col, f.val);
      else if (f.type === "gte") query = query.gte(f.col, f.val);
      else if (f.type === "lt") query = query.lt(f.col, f.val);
      else if (f.type === "lte") query = query.lte(f.col, f.val);
      else if (f.type === "or") query = query.or(f.val);
    }
    return query;
  }

  private async execute(): Promise<{ data: any; error: any; count?: number | null }> {
    const storageKey = `table_${this.table}`;

    // =====================================
    // READ (SELECT)
    // =====================================
    if (this.action === "select") {
      if (navigator.onLine) {
        try {
          let q: any = rawSupabase
            .from(this.table)
            .select(this.selectQuery, this.selectOpts);
          q = this.applyFiltersToSupabaseQuery(q);

          for (const ob of this.orderBys) {
            q = q.order(ob.col, { ascending: ob.ascending });
          }
          if (this.limitCount !== null) q = q.limit(this.limitCount);
          if (this.rangeBounds)
            q = q.range(this.rangeBounds.from, this.rangeBounds.to);
          if (this.isSingle) q = q.single();
          if (this.isMaybeSingle) q = q.maybeSingle();

          const res = await q;
          if (!res.error) {
            if (
              this.filters.length === 0 &&
              !this.isSingle &&
              !this.isMaybeSingle &&
              this.limitCount === null &&
              Array.isArray(res.data)
            ) {
              await set(storageKey, res.data);
            }
            return res;
          }
        } catch {
          // Network failure -> fall back to IndexedDB
        }
      }

      // OFFLINE READ FROM INDEXEDDB
      let localRows: any[] = (await get(storageKey)) || [];
      localRows = localRows.filter((r) => this.matchesFilters(r));
      localRows = await this.attachRelations(localRows);

      for (const ob of this.orderBys) {
        const { col, ascending } = ob;
        localRows.sort((a, b) => {
          const valA = a[col] ?? "";
          const valB = b[col] ?? "";
          if (valA < valB) return ascending ? -1 : 1;
          if (valA > valB) return ascending ? 1 : -1;
          return 0;
        });
      }

      if (this.rangeBounds) {
        localRows = localRows.slice(
          this.rangeBounds.from,
          this.rangeBounds.to + 1
        );
      } else if (this.limitCount !== null) {
        localRows = localRows.slice(0, this.limitCount);
      }

      if (this.isSingle || this.isMaybeSingle) {
        return { data: localRows[0] || null, error: null };
      }
      return { data: localRows, error: null, count: localRows.length };
    }

    // =====================================
    // WRITE (INSERT / UPDATE / DELETE / UPSERT)
    // =====================================
    if (navigator.onLine) {
      try {
        let q: any = rawSupabase.from(this.table);
        if (this.action === "insert") q = q.insert(this.payload).select();
        else if (this.action === "upsert")
          q = q.upsert(this.payload, this.upsertOpts).select();
        else if (this.action === "update") {
          q = this.applyFiltersToSupabaseQuery(q.update(this.payload)).select();
        } else if (this.action === "delete") {
          q = this.applyFiltersToSupabaseQuery(q.delete());
        }

        if (this.isSingle && this.action !== "delete") q = q.single();
        if (this.isMaybeSingle && this.action !== "delete") q = q.maybeSingle();

        const { data, error } = await q;
        if (!error) {
          await this.applyLocalMutation(data);
          return { data, error: null };
        }
      } catch {
        // Offline fallback below
      }
    }

    const mutatedData = await this.applyLocalMutation();
    await addToMutationQueue({
      table: this.table,
      action: this.action,
      payload: this.payload,
      filters: this.filters,
    });

    return {
      data:
        this.isSingle || this.isMaybeSingle
          ? mutatedData[0] || null
          : mutatedData,
      error: null,
    };
  }

  private async applyLocalMutation(serverData?: any): Promise<any[]> {
    const storageKey = `table_${this.table}`;
    let localRows: any[] = (await get(storageKey)) || [];
    let affected: any[] = [];

    if (this.action === "insert" || this.action === "upsert") {
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
            (this.table === "farmers" ||
            this.table === "agents" ||
            this.table === "sub_agents" ||
            this.table === "retailers"
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