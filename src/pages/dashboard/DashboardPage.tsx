import { useEffect, useState, type ComponentType, type ReactNode } from "react";
import { Link, useNavigate } from "react-router-dom";
import { supabase } from "../../lib/supabase";
import {
  FaChartBar,
  FaTractor,
  FaUsers,
  FaCalendarAlt,
  FaMapMarkerAlt,
  FaRoute,
  FaFish,
  FaPaw,
  FaStethoscope,
} from "react-icons/fa";
import { GiChicken, GiRooster } from "react-icons/gi";

function parseArray(field: any): string[] {
  if (!field) return [];
  if (Array.isArray(field)) return field;
  if (typeof field === "string") {
    try {
      const parsed = JSON.parse(field);
      if (Array.isArray(parsed)) return parsed;
    } catch {
      return field
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
    }
  }
  return [];
}
// Visit picker types: one entry per farm category.
type VisitType = {
  key: string;
  label: string;
  path: string;
  Icon: ComponentType<{ className?: string }>;
  cardIconColor: string;
  labelColor: string;
  gradient: string;
  edge: string;
};

type NavCard = {
  title: string;
  path: string;
  icon: ReactNode;
  bg: string;
  edge: string;
  glow: string;
  textColor: string;
  action?: () => void;
};

// Page cache (same pattern as DocTrackingPage): the first visit loads once,
// returning to the page renders instantly from memory with no loading
// spinner. The cache lives for the browser session; a full refresh reloads
// fresh data. The access inputs (role + assigned upazilas) are cached too,
// so the officer scoping stays correct on a return visit.
let cachedPageState: {
  initialized: boolean;
  userId: string;
  role: string;
  assignedUpazilas: string[];
  assignedCategories: string[];
  supplies: any[];
  trackings: any[];
  selectedMonth: string;
  visibleMonthsCount: number;
} | null = null;

export default function Dashboard() {
  const [loading, setLoading] = useState(!cachedPageState?.initialized);
  const [supplies, setSupplies] = useState<any[]>(
    cachedPageState?.supplies || []
  );
  const [trackings, setTrackings] = useState<any[]>(
    cachedPageState?.trackings || []
  );
  const [selectedMonth, setSelectedMonth] = useState<string>(
    cachedPageState?.selectedMonth || "all"
  );
  const [visibleMonthsCount, setVisibleMonthsCount] = useState<number>(
    cachedPageState?.visibleMonthsCount ?? 2
  );
  const [userRole, setUserRole] = useState<string>(
    cachedPageState?.role || "officer"
  );
  const [assignedCategories, setAssignedCategories] = useState<string[]>(
    cachedPageState?.assignedCategories || []
  );
  const [showVisitPicker, setShowVisitPicker] = useState(false);
  const navigate = useNavigate();

  useEffect(() => {
    if (cachedPageState?.initialized) return;
    loadDashboardData();
  }, []);

  async function loadDashboardData() {
    setLoading(true);
    try {
      const {
        data: { user },
      } = await supabase.auth.getUser();

      let uUpazilas: string[] = [];
      let role = "officer";

      if (user) {
        const { data: profile } = await supabase
          .from("users")
          .select("role, upazilas")
          .eq("id", user.id)
          .single();

        if (profile) {
          role = profile.role || "officer";
          uUpazilas = parseArray(profile.upazilas);
        }
      }

      // Assigned categories (Poultry / Cattle / Aqua) drive the Visit
      // card. Fetched separately in a guard so a missing column can
      // never break the profile/role load above.
      let uCategories: string[] = [];
      if (user) {
        try {
          const { data: catData } = await supabase
            .from("users")
            .select("assigned_categories")
            .eq("id", user.id)
            .single();
          uCategories = parseArray(catData?.assigned_categories).map((c) =>
            c === "Fish" ? "Aqua" : c
          );
        } catch {
          uCategories = [];
        }
      }

      const isFullAccess = role === "super_admin" || role === "manager";

      const [{ data: supData }, { data: trkData }] = await Promise.all([
        supabase
          .from("doc_supplies")
          .select("*")
          .order("h_date", { ascending: false }),
        supabase.from("doc_trackings").select("*"),
      ]);

      let finalSupplies: any[] = [];
      let finalTrackings: any[] = [];

      if (isFullAccess || uUpazilas.length === 0) {
        finalSupplies = supData || [];
        finalTrackings = trkData || [];
      } else {
        const { data: agData } = await supabase
          .from("doc_agents")
          .select("doc_id, upazila")
          .in("upazila", uUpazilas);

        const allowedDocIds = new Set((agData || []).map((a) => a.doc_id));
        finalSupplies = (supData || []).filter((s) =>
          allowedDocIds.has(s.doc_id)
        );
        finalTrackings = (trkData || []).filter((t) =>
          allowedDocIds.has(t.doc_id)
        );
      }

      setSupplies(finalSupplies);
      setTrackings(finalTrackings);
      setUserRole(role);
      setAssignedCategories(uCategories);

      cachedPageState = {
        initialized: true,
        userId: user?.id || "",
        role,
        assignedUpazilas: uUpazilas,
        assignedCategories: uCategories,
        supplies: finalSupplies,
        trackings: finalTrackings,
        selectedMonth,
        visibleMonthsCount,
      };
    } catch (err) {
      console.error("Failed to load dashboard data:", err);
    } finally {
      setLoading(false);
    }
  }

  // Group DOC Supplies & Trackings by Month (YYYY-MM)
  const monthlyMap: Record<
    string,
    {
      totalQty: number;
      trackedQty: number;
      untrackedQty: number;
      batchesCount: number;
    }
  > = {};

  supplies.forEach((s) => {
    const ym = String(s.h_date || "").slice(0, 7);
    if (!ym || ym.length !== 7) return;
    if (!monthlyMap[ym]) {
      monthlyMap[ym] = {
        totalQty: 0,
        trackedQty: 0,
        untrackedQty: 0,
        batchesCount: 0,
      };
    }
    monthlyMap[ym].totalQty += Number(s.qty || 0);
    monthlyMap[ym].batchesCount += 1;
  });

  trackings.forEach((t) => {
    const ym = String(t.h_date || "").slice(0, 7);
    if (!ym || !monthlyMap[ym]) return;
    monthlyMap[ym].trackedQty += Number(t.qty || 0);
  });

  Object.keys(monthlyMap).forEach((ym) => {
    const m = monthlyMap[ym];
    m.untrackedQty = Math.max(0, m.totalQty - m.trackedQty);
  });

  const allMonths = Object.keys(monthlyMap).sort().reverse();

  const filteredMonths =
    selectedMonth === "all"
      ? allMonths
      : allMonths.filter((m) => m === selectedMonth);

  const displayedMonths = filteredMonths.slice(0, visibleMonthsCount);

  function formatMonthName(ym: string) {
    const [y, m] = ym.split("-");
    const date = new Date(Number(y), Number(m) - 1, 1);
    return date.toLocaleString("en-US", { month: "short", year: "numeric" });
  }

  // Visit card: category-aware. One category -> its visit page opens
  // directly. Two or three -> a picker shows only those categories.
  // Managers / super admins (and officers with no category set) see
  // all three.
  const VISIT_TYPES: VisitType[] = [
    {
      key: "Poultry",
      label: "Poultry",
      path: "/visits",
      Icon: GiRooster,
      cardIconColor: "text-indigo-600",
      labelColor: "text-indigo-900",
      gradient: "from-indigo-100 to-indigo-200",
      edge: "border-indigo-300",
    },
    {
      key: "Aqua",
      label: "Aqua",
      path: "/aqua-visits",
      Icon: FaFish,
      cardIconColor: "text-sky-600",
      labelColor: "text-sky-900",
      gradient: "from-sky-100 to-sky-200",
      edge: "border-sky-300",
    },
    {
      key: "Cattle",
      label: "Cattle",
      path: "/cattle-visits",
      Icon: FaPaw,
      cardIconColor: "text-amber-600",
      labelColor: "text-amber-900",
      gradient: "from-amber-100 to-amber-200",
      edge: "border-amber-300",
    },
  ];
  const isFullAccess =
    userRole === "super_admin" || userRole === "manager";
  const effectiveVisitTypes =
    isFullAccess || assignedCategories.length === 0
      ? VISIT_TYPES
      : VISIT_TYPES.filter((t) => assignedCategories.includes(t.key));
  const singleVisit =
    effectiveVisitTypes.length === 1 ? effectiveVisitTypes[0] : null;
  const visitCardIcon = singleVisit ? (
    <singleVisit.Icon
      className={`${singleVisit.cardIconColor} text-2xl sm:text-3xl`}
    />
  ) : (
    <FaStethoscope className="text-indigo-600 text-2xl sm:text-3xl" />
  );
  function handleVisitCardClick() {
    if (effectiveVisitTypes.length === 1) {
      navigate(effectiveVisitTypes[0].path);
    } else {
      setShowVisitPicker(true);
    }
  }

  const navCards: NavCard[] = [
    {
      title: "Agents",
      path: "/agents",
      icon: <FaUsers className="text-purple-600 text-2xl sm:text-3xl" />,
      bg: "from-purple-100 to-purple-200",
      edge: "border-b-purple-300",
      glow: "shadow-purple-300/30",
      textColor: "text-purple-900",
    },
    {
      title: "Farmers",
      path: "/farmers",
      icon: <FaTractor className="text-teal-600 text-2xl sm:text-3xl" />,
      bg: "from-teal-100 to-teal-200",
      edge: "border-b-teal-300",
      glow: "shadow-teal-300/30",
      textColor: "text-teal-900",
    },
    {
      title: "Visit",
      path: "/visits",
      icon: visitCardIcon,
      bg: "from-indigo-100 to-indigo-200",
      edge: "border-b-indigo-300",
      glow: "shadow-indigo-300/30",
      textColor: "text-indigo-900",
      action: handleVisitCardClick,
    },
    {
      title: "Reports",
      path: "/reports",
      icon: <FaChartBar className="text-amber-600 text-2xl sm:text-3xl" />,
      bg: "from-amber-100 to-amber-200",
      edge: "border-b-amber-300",
      glow: "shadow-amber-300/30",
      textColor: "text-amber-900",
    },
    {
      title: "Agent Tracking",
      path: "/agent-tracking",
      icon: <FaRoute className="text-rose-600 text-2xl sm:text-3xl" />,
      bg: "from-rose-100 to-rose-200",
      edge: "border-b-rose-300",
      glow: "shadow-rose-300/30",
      textColor: "text-rose-900",
    },
    {
      title: "Farmer Tracking",
      path: "/farmer-tracking",
      icon: <FaMapMarkerAlt className="text-orange-600 text-2xl sm:text-3xl" />,
      bg: "from-orange-100 to-orange-200",
      edge: "border-b-orange-300",
      glow: "shadow-orange-300/30",
      textColor: "text-orange-900",
    },
    {
      title: "DOC Tracking",
      path: "/doc-tracking",
      icon: <GiChicken className="text-emerald-600 text-2xl sm:text-3xl" />,
      bg: "from-emerald-100 to-emerald-200",
      edge: "border-b-emerald-300",
      glow: "shadow-emerald-300/30",
      textColor: "text-emerald-900",
    },
  ];

  // One card for the uniform grid — soft 3D style: light gradient
  // face, soft bottom edge, gentle glow, lift on hover, press on tap.
  // A card with an action renders as a button instead of a link.
  // The wide card (last one) spans 2 columns.
  function renderCard(card: NavCard, wide = false) {
    const content = (
      <>
        <div className="p-4 sm:p-3.5 bg-white rounded-xl shadow-sm shrink-0">
          {card.icon}
        </div>
        <div className="min-w-0 flex-1">
          <h2
            className={`text-xs sm:text-lg font-bold tracking-tight ${card.textColor}`}
          >
            {card.title}
          </h2>
        </div>
      </>
    );
    const cls = `rounded-2xl bg-gradient-to-br ${card.bg} border-b-4 ${card.edge} shadow-lg ${card.glow} p-3 sm:p-5 transition transform hover:-translate-y-1 hover:shadow-xl active:translate-y-0 active:border-b-2 flex flex-col sm:flex-row items-center sm:items-center gap-2 sm:gap-4 text-center sm:text-left ${
      wide ? "col-span-2 " : ""
    }`;
    if (card.action) {
      return (
        <button
          type="button"
          key={card.path}
          onClick={card.action}
          className={`${cls} cursor-pointer`}
        >
          {content}
        </button>
      );
    }
    return (
      <Link key={card.path} to={card.path} className={cls}>
        {content}
      </Link>
    );
  }

  return (
    <div className="p-3 sm:p-6 max-w-7xl mx-auto space-y-5">
      {/* One uniform grid for all cards: 2 per row on mobile, 4 per
          row on PC. The last card (DOC Tracking) spans 2 columns so
          every row stays full. */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-5">
        {navCards.map((card, i) =>
          renderCard(card, i === navCards.length - 1)
        )}
      </div>

      {/* TOTAL DOC BY MONTHS TABLE FORMAT (COMPACT MOBILE FIT) */}
      <div className="bg-white border border-slate-200 rounded-2xl p-3 sm:p-5 shadow-sm space-y-3">
        <div className="flex items-center justify-between gap-2 border-b pb-2.5">
          <div className="flex items-center gap-2">
            <FaCalendarAlt className="text-emerald-600 text-base sm:text-lg" />
            <h3 className="text-sm sm:text-lg font-bold text-slate-800">
              Total DOC by Months
            </h3>
          </div>

          <select
            value={selectedMonth}
            onChange={(e) => {
              setSelectedMonth(e.target.value);
              setVisibleMonthsCount(2);
            }}
            className="border border-slate-300 rounded-lg px-2.5 py-1 text-[11px] sm:text-xs bg-white outline-none focus:border-emerald-500 font-semibold text-slate-700"
          >
            <option value="all">All Months</option>
            {allMonths.map((ym) => (
              <option key={ym} value={ym}>
                {formatMonthName(ym)}
              </option>
            ))}
          </select>
        </div>

        {loading ? (
          <div className="py-8 text-center text-xs text-slate-500">
            Loading monthly DOC data...
          </div>
        ) : displayedMonths.length === 0 ? (
          <div className="py-8 text-center text-xs text-slate-400">
            No monthly DOC records found.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-left border-collapse text-[11px] sm:text-sm">
              <thead>
                <tr className="border-b border-slate-200 text-slate-500 bg-slate-50">
                  <th className="py-2 px-2 sm:px-3 font-bold">Month</th>
                  <th className="py-2 px-2 sm:px-3 font-bold text-right">
                    Total DOC
                  </th>
                  <th className="py-2 px-2 sm:px-3 font-bold text-right">
                    Tracked
                  </th>
                  <th className="py-2 px-2 sm:px-3 font-bold text-right">
                    Untracked
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {displayedMonths.map((ym) => {
                  const data = monthlyMap[ym];
                  const trkPct =
                    data.totalQty > 0
                      ? ((data.trackedQty / data.totalQty) * 100).toFixed(1)
                      : "0.0";
                  const untrkPct =
                    data.totalQty > 0
                      ? ((data.untrackedQty / data.totalQty) * 100).toFixed(1)
                      : "0.0";

                  return (
                    <tr key={ym} className="hover:bg-slate-50/50">
                      <td className="py-2.5 px-2 sm:px-3 font-bold text-slate-800 whitespace-nowrap">
                        {formatMonthName(ym)}
                      </td>
                      <td className="py-2.5 px-2 sm:px-3 text-right font-extrabold text-slate-900 whitespace-nowrap">
                        {data.totalQty.toLocaleString()}
                      </td>
                      <td className="py-2.5 px-2 sm:px-3 text-right font-bold text-emerald-600 whitespace-nowrap">
                        {data.trackedQty.toLocaleString()} ({trkPct}%)
                      </td>
                      <td className="py-2.5 px-2 sm:px-3 text-right font-bold text-amber-600 whitespace-nowrap">
                        {data.untrackedQty.toLocaleString()} ({untrkPct}%)
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}

        {/* Load More Button (Shows 2 initially, then adds 5 more) */}
        {!loading &&
          selectedMonth === "all" &&
          visibleMonthsCount < allMonths.length && (
            <div className="text-center pt-1">
              <button
                type="button"
                onClick={() => setVisibleMonthsCount((prev) => prev + 5)}
                className="px-4 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold shadow-sm transition"
              >
                Load More ({allMonths.length - visibleMonthsCount} remaining)
              </button>
            </div>
          )}
      </div>
      {/* Visit picker: only the user's categories, as 3D buttons. */}
      {showVisitPicker && (
        <div
          className="fixed inset-0 z-[100] bg-black/50 backdrop-blur-sm flex items-center justify-center p-4"
          onClick={() => setShowVisitPicker(false)}
        >
          <div
            className="bg-white rounded-3xl shadow-2xl w-full max-w-sm p-5"
            onClick={(e) => e.stopPropagation()}
          >
            <h3 className="text-base font-bold text-slate-800 text-center">
              Choose Visit
            </h3>
            <p className="text-xs text-slate-500 text-center mt-1 mb-4">
              Select which visit to open
            </p>
            <div
              className={`grid gap-3 ${
                effectiveVisitTypes.length === 2
                  ? "grid-cols-2"
                  : "grid-cols-3"
              }`}
            >
              {effectiveVisitTypes.map((t) => (
                <Link
                  key={t.key}
                  to={t.path}
                  onClick={() => setShowVisitPicker(false)}
                  className={`flex flex-col items-center gap-2 rounded-2xl bg-gradient-to-b ${t.gradient} py-4 px-2 shadow-lg border-b-4 ${t.edge} transition hover:-translate-y-1 hover:shadow-xl active:translate-y-0 active:border-b-2`}
                >
                  <t.Icon className={`text-3xl ${t.cardIconColor}`} />
                  <span className={`text-xs font-bold ${t.labelColor}`}>{t.label}</span>
                </Link>
              ))}
            </div>
            <button
              type="button"
              onClick={() => setShowVisitPicker(false)}
              className="mt-4 w-full py-2 rounded-xl border border-slate-300 text-slate-700 text-xs font-semibold"
            >
              Cancel
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
