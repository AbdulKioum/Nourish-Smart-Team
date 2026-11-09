import { Outlet, Navigate } from "react-router-dom";
import { useState, useRef } from "react";
import { useAuth } from "../context/AuthContext";

import Sidebar from "../components/common/Sidebar";
import Navbar from "../components/common/Navbar";

// একবার ওয়ার্কস্পেস লোড হয়ে গেলে মিনিমাইজ বা ট্যাব সুইচে আর লোডিং স্ক্রিন দেখাবে না
let workspaceInitialized = false;

// প্রিমিয়াম সিম্পল লোডার কম্পোনেন্ট (ক্রিম/অফ-হোয়াইট ব্যাকগ্রাউন্ড ও লিনিয়ার মুভিং প্রোগ্রেস বার সহ)
function SimpleLoader() {
  return (
    <div className="fixed inset-0 z-50 bg-[#F9F8F6] flex flex-col items-center justify-center space-y-4">
      {/* লোগো ও হালকা পাল্সিং গ্লো */}
      <div className="relative flex items-center justify-center">
        <div className="absolute w-20 h-20 rounded-full bg-emerald-500/10 animate-ping" />
        <img
          src="/logo smart team.png"
          alt="Loading..."
          className="w-14 h-14 object-contain relative z-10 drop-shadow-sm"
        />
      </div>

      {/* ব্র্যান্ডিং ও বাম থেকে ডানে মুভ করা স্মুথ লাইন অ্যানিমেশন */}
      <div className="flex flex-col items-center space-y-2 bg-white px-5 py-3 rounded-2xl shadow-sm border border-slate-200/80 w-52">
        <span className="text-xs font-bold text-slate-700 tracking-wide">Preparing workspace...</span>
        
        {/* প্রোগ্রেস বার কন্টেইনার */}
        <div className="w-full h-1.5 bg-slate-100 rounded-full overflow-hidden relative">
          <div className="absolute top-0 bottom-0 w-1/2 bg-emerald-600 rounded-full animate-[shimmer_1.5s_infinite_linear]" />
        </div>
      </div>

      {/* অ্যানিমেশনের জন্য কাস্টম স্টাইল */}
      <style>{`
        @keyframes shimmer {
          0% { left: -60%; }
          100% { left: 100%; }
        }
      `}</style>
    </div>
  );
}

export default function MainLayout() {
  const { role, userName, email, loading } = useAuth();
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const lastKnownEmailRef = useRef<string | null>(email || null);

  if (email) {
    workspaceInitialized = true;
    lastKnownEmailRef.current = email;
  }

  // শুধুমাত্র প্রথমবার অ্যাপ ওপেন করার সময় এই প্রিমিয়াম লোডার স্ক্রিন দেখাবে
  if (loading && !workspaceInitialized) {
    return <SimpleLoader />;
  }

  // ব্যাকগ্রাউন্ড রিফ্রেশের সময় যদি সাময়িক loading থাকে কিন্তু আগে ইমেইল ছিল, তাহলে পেজ আনমাউন্ট করবে না
  const activeEmail = email || (loading ? lastKnownEmailRef.current : null);

  if (!activeEmail) {
    workspaceInitialized = false;
    lastKnownEmailRef.current = null;
    return <Navigate to="/login" replace />;
  }

  return (
    <div className="flex h-screen overflow-hidden bg-slate-50">
      {/* Slide-Over Sidebar Drawer */}
      <Sidebar
        role={role}
        userName={userName}
        sidebarOpen={sidebarOpen}
        setSidebarOpen={setSidebarOpen}
      />

      {/* Main Content Area */}
      <div className="flex-1 min-w-0 flex flex-col h-full overflow-hidden">
        <Navbar
          role={role}
          userName={userName}
          sidebarOpen={sidebarOpen}
          setSidebarOpen={setSidebarOpen}
        />

        {/* Mobile-friendly responsive padding (p-2 sm:p-4 md:p-6) */}
        <main className="flex-1 overflow-y-auto overflow-x-hidden bg-gradient-to-br from-slate-50 to-slate-100 p-2 sm:p-4 md:p-6">
          <Outlet />
        </main>
      </div>
    </div>
  );
}