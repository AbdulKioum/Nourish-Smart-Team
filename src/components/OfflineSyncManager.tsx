import { useEffect, useState } from "react";
import { get } from "idb-keyval";
import {
  rawSupabase,
  pullAllTablesToLocal,
  processMutationQueue,
  SYNC_TABLES,
} from "../lib/supabase";
import { FaWifi, FaSyncAlt, FaCloudUploadAlt } from "react-icons/fa";

export default function OfflineSyncManager() {
  const [isOnline, setIsOnline] = useState<boolean>(navigator.onLine);
  const [pendingCount, setPendingCount] = useState<number>(0);
  const [syncing, setSyncing] = useState<boolean>(false);

  async function updateQueueCount() {
    const q: any[] = (await get("fams_offline_mutation_queue")) || [];
    setPendingCount(q.length);
  }

  async function runFullSync() {
    if (!navigator.onLine) return;
    setSyncing(true);
    try {
      await processMutationQueue();
      await pullAllTablesToLocal();
      await updateQueueCount();
    } finally {
      setSyncing(false);
    }
  }

  useEffect(() => {
    updateQueueCount();
    if (navigator.onLine) {
      runFullSync();
    }

    const handleOnline = () => {
      setIsOnline(true);
      runFullSync();
    };

    const handleOffline = () => {
      setIsOnline(false);
    };

    const handleQueueChange = () => {
      updateQueueCount();
      if (navigator.onLine) {
        processMutationQueue();
      }
    };

    window.addEventListener("online", handleOnline);
    window.addEventListener("offline", handleOffline);
    window.addEventListener("fams-queue-updated", handleQueueChange);

    // Background Realtime Listener (Onno user update korle auto sync hobe)
    const channel = rawSupabase
      .channel("fams_global_realtime_sync")
      .on(
        "postgres_changes",
        { event: "*", schema: "public" },
        (payload) => {
          if (SYNC_TABLES.includes(payload.table)) {
            pullAllTablesToLocal();
          }
        }
      )
      .subscribe();

    // Every 2 minutes background check if online
    const interval = setInterval(() => {
      if (navigator.onLine) {
        runFullSync();
      }
    }, 120000);

    return () => {
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
      window.removeEventListener("fams-queue-updated", handleQueueChange);
      rawSupabase.removeChannel(channel);
      clearInterval(interval);
    };
  }, []);

  if (isOnline && pendingCount === 0 && !syncing) return null;

  return (
    <div className="fixed bottom-3 right-3 z-50 flex items-center gap-2 px-3.5 py-2 rounded-full shadow-lg border text-xs font-bold transition-all bg-slate-900 text-white border-slate-700">
      {!isOnline ? (
        <>
          <span className="h-2 w-2 rounded-full bg-amber-400 animate-ping" />
          <FaWifi className="text-amber-400" />
          <span>Offline Mode</span>
          {pendingCount > 0 && (
            <span className="bg-amber-500 text-slate-950 px-2 py-0.5 rounded-full text-[10px]">
              {pendingCount} Pending Sync
            </span>
          )}
        </>
      ) : syncing || pendingCount > 0 ? (
        <>
          <FaSyncAlt className="animate-spin text-emerald-400" />
          <span>Syncing Data ({pendingCount})...</span>
        </>
      ) : (
        <>
          <FaCloudUploadAlt className="text-emerald-400" />
          <span>Synced</span>
        </>
      )}
    </div>
  );
}