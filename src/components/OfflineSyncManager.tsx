import { useEffect, useState, useRef } from "react";
import { get } from "idb-keyval";
import { App as CapacitorApp } from "@capacitor/app";
import {
  rawSupabase,
  pullAllTablesToLocal,
  processMutationQueue,
  SYNC_TABLES,
} from "../lib/supabase";
import { FaWifi, FaSyncAlt, FaCloudUploadAlt, FaSignOutAlt } from "react-icons/fa";

export default function OfflineSyncManager() {
  const [isOnline, setIsOnline] = useState<boolean>(navigator.onLine);
  const [pendingCount, setPendingCount] = useState<number>(0);
  const [syncing, setSyncing] = useState<boolean>(false);
  const [showExitModal, setShowExitModal] = useState<boolean>(false);

  const exitModalOpenRef = useRef(false);
  useEffect(() => {
    exitModalOpenRef.current = showExitModal;
  }, [showExitModal]);

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
    // 1. Android Hardware Back Button Handler with Exit Confirmation
    const backButtonListener = CapacitorApp.addListener(
      "backButton",
      ({ canGoBack }) => {
        // Jodi exit warning popup open thake, tobe back chaple popup close hobe
        if (exitModalOpenRef.current) {
          setShowExitModal(false);
          return;
        }

        const currentPath = window.location.pathname;
        const isRootPage =
          currentPath === "/" ||
          currentPath === "/dashboard" ||
          currentPath === "/login";

        if (!isRootPage && (canGoBack || window.history.length > 1)) {
          window.history.back();
        } else {
          // Home ba Login page-e thakle warning popup dekhabe
          setShowExitModal(true);
        }
      }
    );

    // 2. Offline & Realtime Sync Setup
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

    const interval = setInterval(() => {
      if (navigator.onLine) {
        runFullSync();
      }
    }, 120000);

    return () => {
      backButtonListener.then((listener) => listener.remove());
      window.removeEventListener("online", handleOnline);
      window.removeEventListener("offline", handleOffline);
      window.removeEventListener("fams-queue-updated", handleQueueChange);
      rawSupabase.removeChannel(channel);
      clearInterval(interval);
    };
  }, []);

  return (
    <>
      {/* Offline / Syncing Status Badge */}
      {(!isOnline || pendingCount > 0 || syncing) && (
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
      )}

      {/* Exit App Confirmation Modal */}
      {showExitModal && (
        <div className="fixed inset-0 z-[9999] bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-xs p-5 border border-slate-200 text-center space-y-4">
            <div className="w-12 h-12 rounded-full bg-red-50 text-red-600 flex items-center justify-center mx-auto">
              <FaSignOutAlt size={22} />
            </div>

            <div className="space-y-1">
              <h3 className="text-base font-bold text-slate-800">
                Exit Application?
              </h3>
              <p className="text-xs text-slate-500">
                Are you sure you want to close Smart Team app?
              </p>
            </div>

            <div className="flex items-center gap-2.5 pt-1">
              <button
                type="button"
                onClick={() => setShowExitModal(false)}
                className="flex-1 py-2 rounded-xl border border-slate-300 text-slate-700 hover:bg-slate-100 text-xs font-bold transition"
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={() => {
                  setShowExitModal(false);
                  CapacitorApp.exitApp();
                }}
                className="flex-1 py-2 rounded-xl bg-red-600 hover:bg-red-700 text-white text-xs font-bold shadow-sm transition"
              >
                Exit App
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}