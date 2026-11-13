import React, { useEffect, useState } from "react";
import { Link, useLocation } from "react-router-dom";
import {
  FaHome,
  FaBuilding,
  FaUsers,
  FaSitemap,
  FaStore,
  FaUserFriends,
  FaClipboardList,
  FaChartBar,
  FaUserShield,
  FaTimes,
  FaSignOutAlt,
  FaUserCircle,
  FaMapMarkerAlt,
  FaStethoscope,
  FaDownload,
} from "react-icons/fa";
import { Capacitor } from "@capacitor/core";
import { supabase } from "../../lib/supabase";
import {
  compressImageToTinyBase64,
  UserAvatar,
} from "../../utils/avatarHelper";


const APK_DOWNLOAD_URL =
  "https://vqqlckzchnppvidkckrw.supabase.co/storage/v1/object/public/app-releases/Nourish-Smart-Team.apk";

function parseList(val: any): string[] {
  if (!val) return [];
  if (Array.isArray(val))
    return val.map((s) => String(s).trim()).filter(Boolean);
  if (typeof val === "string") {
    try {
      const parsed = JSON.parse(val);
      if (Array.isArray(parsed)) {
        return parsed.map((s) => String(s).trim()).filter(Boolean);
      }
    } catch {
      return val
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
    }
  }
  return [];
}

export default function Sidebar({
  role,
  userName,
  sidebarOpen,
  setSidebarOpen,
}: {
  role: string;
  userName?: string;
  sidebarOpen: boolean;
  setSidebarOpen: React.Dispatch<React.SetStateAction<boolean>>;
}) {
  const location = useLocation();

  
  const isWebBrowser = !Capacitor.isNativePlatform();

  const [showProfileModal, setShowProfileModal] = useState(false);
  const [savingMeta, setSavingMeta] = useState(false);
  const [profileDetails, setProfileDetails] = useState({
    id: "",
    fullName: userName || "User",
    email: "",
    role: role || "officer",
    designation: "Customer Service Officer",
    region: "Mymensingh Zone",
    avatar_url: "",
    districts: [] as string[],
    upazilas: [] as string[],
  });

  useEffect(() => {
    async function loadUserProfile() {
      const {
        data: { user },
      } = await supabase.auth.getUser();
      if (!user) return;

      const { data } = await supabase
        .from("users")
        .select(
          "full_name, email, role, designation, region, avatar_url, districts, upazilas"
        )
        .eq("id", user.id)
        .single();

      setProfileDetails({
        id: user.id,
        fullName:
          data?.full_name ||
          userName ||
          user.email?.split("@")[0] ||
          "User",
        email: data?.email || user.email || "",
        role: data?.role || role || "officer",
        designation: data?.designation || "CSO",
        region: data?.region || "Mymensingh Region",
        avatar_url: data?.avatar_url || "",
        districts: parseList(data?.districts),
        upazilas: parseList(data?.upazilas),
      });
    }

    loadUserProfile();
  }, [userName, role]);

  async function handleProfilePhotoChange(
    e: React.ChangeEvent<HTMLInputElement>
  ) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const tinyBase64 = await compressImageToTinyBase64(file, 80);
      setProfileDetails((prev) => ({ ...prev, avatar_url: tinyBase64 }));
    } catch (err) {
      console.error("Image compress error:", err);
    }
  }

  async function handleSaveProfileMeta() {
    if (!profileDetails.id) return;
    setSavingMeta(true);
    await supabase
      .from("users")
      .update({
        designation: profileDetails.designation,
        region: profileDetails.region,
        avatar_url: profileDetails.avatar_url || null,
      })
      .eq("id", profileDetails.id);
    setSavingMeta(false);
    setShowProfileModal(false);
  }

  async function handleLogout() {
    await supabase.auth.signOut();
    window.location.href = "/login";
  }

  function handleMenuClick() {
    setSidebarOpen(false);
  }

  const menuClass = (path: string) =>
    `flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-sm font-medium transition-all duration-200 ${
      location.pathname === path
        ? "bg-emerald-600 text-white shadow-sm"
        : "text-slate-700 hover:bg-emerald-50 hover:text-emerald-700"
    }`;

  return (
    <>
      {sidebarOpen && (
        <div
          onClick={() => setSidebarOpen(false)}
          className="fixed inset-0 z-[9998] bg-black/40 backdrop-blur-[1px] transition-opacity"
        />
      )}

      <aside
        className={`fixed top-0 left-0 z-[9999] h-full w-64 max-w-[80vw] bg-white border-r border-slate-200 shadow-2xl flex flex-col transform transition-transform duration-300 ease-in-out ${
          sidebarOpen ? "translate-x-0" : "-translate-x-full"
        }`}
      >
        <div className="flex items-center justify-between px-4 h-14 sm:h-16 border-b border-slate-100 shrink-0">
          <div className="flex items-center gap-2.5">
            <img
              src="/logo smart team.png"
              alt="Logo"
              className="w-9 h-9 object-contain"
            />
            <span className="font-bold text-slate-800 text-base tracking-tight">
              Menu
            </span>
          </div>

          <button
            onClick={() => setSidebarOpen(false)}
            className="w-8 h-8 rounded-lg text-slate-400 hover:text-slate-700 hover:bg-slate-100 flex items-center justify-center transition"
            aria-label="Close Sidebar"
          >
            <FaTimes size={16} />
          </button>
        </div>

        <nav className="flex-1 overflow-y-auto flex flex-col gap-1 px-3 py-3">
          <Link to="/" onClick={handleMenuClick} className={menuClass("/")}>
            <FaHome className="text-base shrink-0" />
            <span>Dashboard</span>
          </Link>

          <div className="mt-3 mb-1 px-3 text-[11px] font-bold uppercase tracking-wider text-slate-400">
            Visit
          </div>

          <Link
            to="/visits"
            onClick={handleMenuClick}
            className={menuClass("/visits")}
          >
            <FaStethoscope className="text-base shrink-0" />
            <span>Vet Visits</span>
          </Link>

          <Link
            to="/reports"
            onClick={handleMenuClick}
            className={menuClass("/reports")}
          >
            <FaChartBar className="text-base shrink-0" />
            <span>Reports</span>
          </Link>

          <div className="mt-3 mb-1 px-3 text-[11px] font-bold uppercase tracking-wider text-slate-400">
            Tracking
          </div>

          <Link
            to="/doc-tracking"
            onClick={handleMenuClick}
            className={menuClass("/doc-tracking")}
          >
            <FaClipboardList className="text-base shrink-0" />
            <span>DOC Tracking</span>
          </Link>

          <Link
            to="/farmer-tracking"
            onClick={handleMenuClick}
            className={menuClass("/farmer-tracking")}
          >
            <FaMapMarkerAlt className="text-base shrink-0" />
            <span>Farmer Tracking</span>
          </Link>

          <Link
            to="/agent-tracking"
            onClick={handleMenuClick}
            className={menuClass("/agent-tracking")}
          >
            <FaClipboardList className="text-base shrink-0" />
            <span>Agent Tracking</span>
          </Link>

          <div className="mt-3 mb-1 px-3 text-[11px] font-bold uppercase tracking-wider text-slate-400">
            Management
          </div>

          <Link
            to="/farmers"
            onClick={handleMenuClick}
            className={menuClass("/farmers")}
          >
            <FaUserFriends className="text-base shrink-0" />
            <span>Farmers</span>
          </Link>

          <Link
            to="/agents"
            onClick={handleMenuClick}
            className={menuClass("/agents")}
          >
            <FaUsers className="text-base shrink-0" />
            <span>Agents</span>
          </Link>

          <Link
            to="/doc-agents"
            onClick={handleMenuClick}
            className={menuClass("/doc-agents")}
          >
            <FaUsers className="text-base shrink-0" />
            <span>DOC Agents</span>
          </Link>

          <Link
            to="/sub-agents"
            onClick={handleMenuClick}
            className={menuClass("/sub-agents")}
          >
            <FaSitemap className="text-base shrink-0" />
            <span>Sub Agents</span>
          </Link>

          <Link
            to="/retailers"
            onClick={handleMenuClick}
            className={menuClass("/retailers")}
          >
            <FaStore className="text-base shrink-0" />
            <span>Retailers</span>
          </Link>

          <Link
            to="/companies"
            onClick={handleMenuClick}
            className={menuClass("/companies")}
          >
            <FaBuilding className="text-base shrink-0" />
            <span>Companies</span>
          </Link>

          {role === "super_admin" && (
            <>
              <div className="mt-3 mb-1 px-3 text-[11px] font-bold uppercase tracking-wider text-slate-400">
                Administration
              </div>

              <Link
                to="/users"
                onClick={handleMenuClick}
                className={menuClass("/users")}
              >
                <FaUserShield className="text-base shrink-0" />
                <span>Users</span>
              </Link>
            </>
          )}
        </nav>

        <div className="mt-auto border-t border-slate-200 p-3.5 bg-slate-50/70 space-y-2.5 shrink-0">
          
          {isWebBrowser && (
            <a
              href={APK_DOWNLOAD_URL}
              download
              target="_blank"
              rel="noopener noreferrer"
              onClick={handleMenuClick}
              className="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 active:scale-[0.99] text-white py-2 px-3 rounded-xl text-xs sm:text-sm font-semibold shadow-sm transition"
            >
              <FaDownload className="text-xs shrink-0" />
              <span>Download App</span>
            </a>
          )}

          <button
            type="button"
            onClick={() => setShowProfileModal(true)}
            title="Click to view profile & assigned area"
            className="w-full flex items-center gap-2.5 px-2 py-1.5 rounded-xl hover:bg-emerald-50/80 transition text-left"
          >
            <UserAvatar
              name={profileDetails.fullName || userName || "User"}
              avatarUrl={profileDetails.avatar_url}
              size="w-8 h-8"
              textSize="text-xs"
            />
            <div className="min-w-0 flex-1">
              <div className="font-semibold text-sm text-slate-800 truncate">
                {profileDetails.fullName || userName || "User"}
              </div>
              <div className="text-xs text-emerald-700 font-medium capitalize truncate">
                {role ? role.replaceAll("_", " ") : "Officer"} • View Profile
              </div>
            </div>
          </button>

          <button
            onClick={handleLogout}
            className="w-full flex items-center justify-center gap-2 bg-red-500 hover:bg-red-600 active:scale-[0.99] text-white py-2 px-3 rounded-xl text-xs sm:text-sm font-medium shadow-sm transition"
          >
            <FaSignOutAlt />
            <span>Logout</span>
          </button>
        </div>
      </aside>

      {/* Profile Details Popup Modal */}
      {showProfileModal && (
        <div className="fixed inset-0 z-[10000] bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-sm overflow-hidden border border-slate-200">
            <div className="flex items-center justify-between px-4 py-3 bg-slate-100 border-b border-slate-200">
              <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                <FaUserCircle className="text-emerald-600 text-base" />
                Profile Details
              </h3>
              <button
                onClick={() => setShowProfileModal(false)}
                className="w-7 h-7 rounded-lg hover:bg-slate-200 flex items-center justify-center text-slate-500"
              >
                <FaTimes size={13} />
              </button>
            </div>

            <div className="p-4 space-y-3 text-xs">
              <div className="flex items-center gap-3 pb-3 border-b border-slate-100">
                <div className="flex flex-col items-center gap-1 shrink-0">
                  <UserAvatar
                    name={profileDetails.fullName || "U"}
                    avatarUrl={profileDetails.avatar_url}
                    size="w-12 h-12"
                    textSize="text-base"
                  />
                  <label className="cursor-pointer text-[10px] font-bold text-emerald-700 hover:text-emerald-800 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 px-2 py-0.5 rounded transition">
                    {profileDetails.avatar_url ? "Change" : "Add Photo"}
                    <input
                      type="file"
                      accept="image/*"
                      onChange={handleProfilePhotoChange}
                      className="hidden"
                    />
                  </label>
                </div>

                <div className="min-w-0 flex-1">
                  <div className="font-bold text-sm text-slate-800 truncate">
                    {profileDetails.fullName}
                  </div>
                  <div className="text-slate-500 truncate">
                    {profileDetails.email || "-"}
                  </div>
                  <span className="inline-block mt-1 px-2 py-0.5 rounded-full bg-emerald-50 border border-emerald-200 text-emerald-700 text-[10px] font-semibold uppercase">
                    {(profileDetails.role || "officer").replaceAll("_", " ")}
                  </span>
                </div>
              </div>

              <div className="grid grid-cols-1 gap-2">
                <div>
                  <label className="text-slate-500 font-semibold uppercase text-[10px] block mb-1">
                    Designation
                  </label>
                  <input
                    type="text"
                    value={profileDetails.designation}
                    onChange={(e) =>
                      setProfileDetails({
                        ...profileDetails,
                        designation: e.target.value,
                      })
                    }
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs outline-none focus:border-emerald-500"
                  />
                </div>

                <div>
                  <label className="text-slate-500 font-semibold uppercase text-[10px] block mb-1">
                    Region / Zone
                  </label>
                  <input
                    type="text"
                    value={profileDetails.region}
                    onChange={(e) =>
                      setProfileDetails({
                        ...profileDetails,
                        region: e.target.value,
                      })
                    }
                    className="w-full border border-slate-300 rounded-lg px-2.5 py-1.5 text-xs outline-none focus:border-emerald-500"
                  />
                </div>
              </div>

              <div>
                <div className="text-slate-400 font-semibold uppercase text-[10px] mb-1">
                  Assigned Districts
                </div>
                <div className="flex flex-wrap gap-1">
                  {profileDetails.districts.length === 0 ? (
                    <span className="text-slate-600 font-medium">
                      All Districts
                    </span>
                  ) : (
                    profileDetails.districts.map((d) => (
                      <span
                        key={d}
                        className="px-2 py-0.5 rounded bg-slate-100 border border-slate-200 text-slate-700 font-medium"
                      >
                        {d}
                      </span>
                    ))
                  )}
                </div>
              </div>

              <div>
                <div className="text-slate-400 font-semibold uppercase text-[10px] mb-1">
                  Assigned Upazilas
                </div>
                <div className="flex flex-wrap gap-1">
                  {profileDetails.upazilas.length === 0 ? (
                    <span className="text-slate-600 font-medium">
                      All Upazilas
                    </span>
                  ) : (
                    profileDetails.upazilas.map((u) => (
                      <span
                        key={u}
                        className="px-2 py-0.5 rounded bg-emerald-50 border border-emerald-200 text-emerald-800 font-semibold"
                      >
                        {u}
                      </span>
                    ))
                  )}
                </div>
              </div>
            </div>

            <div className="px-4 py-2.5 bg-slate-50 border-t border-slate-100 flex justify-end gap-2">
              <button
                onClick={() => setShowProfileModal(false)}
                className="px-3.5 py-1.5 rounded-lg border border-slate-300 text-slate-700 text-xs font-semibold"
              >
                Close
              </button>
              <button
                onClick={handleSaveProfileMeta}
                disabled={savingMeta}
                className="px-4 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold"
              >
                {savingMeta ? "Saving..." : "Save"}
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}