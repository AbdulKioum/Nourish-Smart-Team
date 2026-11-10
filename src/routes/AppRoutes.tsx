import {
  BrowserRouter,
  Routes,
  Route,
  Navigate,
} from "react-router-dom";
import { useEffect, useState, useRef } from "react";
import { PushNotifications } from "@capacitor/push-notifications";
import { Capacitor } from "@capacitor/core";
import { supabase } from "../lib/supabase";

import MainLayout from "../layouts/MainLayout";
import LoginPage from "../pages/auth/LoginPage";

import DashboardPage from "../pages/dashboard/DashboardPage";
import CompaniesPage from "../pages/companies/CompaniesPage";
import AgentsPage from "../pages/agents/AgentsPage";
import SubAgentPage from "../pages/sub-agents/SubAgentPage";
import RetailersPage from "../pages/retailers/RetailerPage";
import FarmersPage from "../pages/farmers/FarmersPage";
import VisitsPage from "../pages/visits/VisitsPage";
import ReportsPage from "../pages/reports/ReportsPage";
import UsersPage from "../pages/users/UsersPage";
import FarmerStatusPage from "../pages/farmers/FarmerStatusPage";
import DocAgentsPage from "../pages/doc/DocAgentsPage";
import DocTrackingPage from "../pages/doc/DocTrackingPage";
import FarmerTrackingPage from "../pages/farmers/FarmerTrackingPage"; 

import {
  setupGlobalNativeGeolocation,
  setupGlobalNativeFileDownload,
} from "../utils/nativeAppHelper";

setupGlobalNativeGeolocation();
setupGlobalNativeFileDownload();

// Loader
function SimpleLoader() {
  return (
    <div className="fixed inset-0 z-50 bg-[#F9F8F6] flex flex-col items-center justify-center space-y-4">
      <div className="relative flex items-center justify-center">
        <div className="absolute w-20 h-20 rounded-full bg-emerald-500/10 animate-ping" />
        <img
          src="/logo smart team.png"
          alt="Loading..."
          className="w-14 h-14 object-contain relative z-10 drop-shadow-sm"
        />
      </div>

      <div className="flex flex-col items-center space-y-2 bg-white px-5 py-3 rounded-2xl shadow-sm border border-slate-200/80 w-48">
        <span className="text-xs font-bold text-slate-700 tracking-wide">Loading Smart Team...</span>
        <div className="w-full h-1.5 bg-slate-100 rounded-full overflow-hidden relative">
          <div className="absolute top-0 bottom-0 w-1/2 bg-emerald-600 rounded-full animate-[shimmer_1.5s_infinite_linear]" />
        </div>
      </div>

      <style>{`
        @keyframes shimmer {
          0% { left: -60%; }
          100% { left: 100%; }
        }
      `}</style>
    </div>
  );
}

export default function AppRoutes() {
  const [loading, setLoading] = useState(true);
  const [session, setSession] = useState<any>(null);
  const currentUserIdRef = useRef<string | null>(null);

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => {
      currentUserIdRef.current = data.session?.user?.id || null;
      setSession(data.session);
      setLoading(false);
    });

    const { data: listener } = supabase.auth.onAuthStateChange(
      (event, newSession) => {
        const newUserId = newSession?.user?.id || null;

        if (event === "SIGNED_OUT") {
          currentUserIdRef.current = null;
          setSession(null);
        } else if (newUserId !== currentUserIdRef.current) {
          currentUserIdRef.current = newUserId;
          setSession(newSession);
        }
      }
    );

    return () => listener.subscription.unsubscribe();
  }, []);

  // Push Notification Setup when session/user is available
  useEffect(() => {
    const userId = session?.user?.id;
    const userEmail = session?.user?.email;

    // Run only on Native Android/iOS platform and when user is logged in
    if (!Capacitor.isNativePlatform() || !userId) return;

    const initPush = async () => {
      try {
        await PushNotifications.removeAllListeners();

        // 1. Listen for FCM Token registration and save to 'users' table
        await PushNotifications.addListener("registration", async (token) => {
          console.log("FCM Token:", token.value);

          // Try updating by user's Auth UUID (id)
          const { data, error } = await supabase
            .from("users")
            .update({ fcm_token: token.value })
            .eq("id", userId)
            .select();

          if (error) {
            alert("DB Update Error: " + error.message);
            return;
          }

          // Fallback: If 'users.id' does not match Auth UUID, try matching by email
          if ((!data || data.length === 0) && userEmail) {
            await supabase
              .from("users")
              .update({ fcm_token: token.value })
              .eq("email", userEmail);
          }
        });

        // 2. Listen for registration errors
        await PushNotifications.addListener("registrationError", (err) => {
          alert("FCM Registration Error: " + JSON.stringify(err));
        });

        // 3. Handle notification click action
        await PushNotifications.addListener(
          "pushNotificationActionPerformed",
          (notification) => {
            console.log("Push action performed:", notification);
          }
        );

        // 4. Check & Request permission, then register
        let permStatus = await PushNotifications.checkPermissions();
        if (permStatus.receive === "prompt") {
          permStatus = await PushNotifications.requestPermissions();
        }

        if (permStatus.receive === "granted") {
          await PushNotifications.register();
        }
      } catch (err: any) {
        console.error("Push setup failed:", err);
      }
    };

    initPush();
  }, [session?.user?.id]);

  if (loading) {
    return <SimpleLoader />;
  }

  return (
    <BrowserRouter>
      <Routes>
        <Route path="/login" element={<LoginPage />} />

        {/* Authenticated Layout Routes */}
        <Route
          element={
            session ? (
              <MainLayout />
            ) : (
              <Navigate to="/login" replace />
            )
          }
        >
          <Route path="/" element={<DashboardPage />} />
          <Route path="/companies" element={<CompaniesPage />} />
          <Route path="/agents" element={<AgentsPage />} />
          <Route path="/sub-agents" element={<SubAgentPage />} />
          <Route path="/retailers" element={<RetailersPage />} />
          <Route path="/farmers" element={<FarmersPage />} />
          <Route path="/visits" element={<VisitsPage />} />
          <Route path="/reports" element={<ReportsPage />} />
          <Route path="/users" element={<UsersPage />} />
          <Route path="/farmer-status" element={<FarmerStatusPage />} />
          <Route path="/doc-agents" element={<DocAgentsPage />} />
          <Route path="/doc-tracking" element={<DocTrackingPage />} />
          <Route path="/farmer-tracking" element={<FarmerTrackingPage />} />
        </Route>
      </Routes>
    </BrowserRouter>
  );
}