import { useEffect, useState } from "react";
import { Link } from "react-router-dom";
import { supabase } from "../../lib/supabase";
import {
  FaClipboardList,
  FaStethoscope,
  FaChartBar,
  FaUserFriends,
  FaUsers,
  FaCalendarAlt,
} from "react-icons/fa";

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

export default function Dashboard() {
  const [loading, setLoading] = useState(true);
  const [supplies, setSupplies] = useState<any[]>([]);
  const [trackings, setTrackings] = useState<any[]>([]);
  const [selectedMonth, setSelectedMonth] = useState<string>("all");
  const [visibleMonthsCount, setVisibleMonthsCount] = useState<number>(2);

  useEffect(() => {
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

      const isFullAccess = role === "super_admin" || role === "manager";

      const [{ data: supData }, { data: trkData }] = await Promise.all([
        supabase
          .from("doc_supplies")
          .select("*")
          .order("h_date", { ascending: false }),
        supabase.from("doc_trackings").select("*"),
      ]);

      if (isFullAccess || uUpazilas.length === 0) {
        setSupplies(supData || []);
        setTrackings(trkData || []);
      } else {
        const { data: agData } = await supabase
          .from("doc_agents")
          .select("doc_id, upazila")
          .in("upazila", uUpazilas);

        const allowedDocIds = new Set((agData || []).map((a) => a.doc_id));
        const filteredSup = (supData || []).filter((s) =>
          allowedDocIds.has(s.doc_id)
        );
        const filteredTrk = (trkData || []).filter((t) =>
          allowedDocIds.has(t.doc_id)
        );

        setSupplies(filteredSup);
        setTrackings(filteredTrk);
      }
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

  const navCards = [
    {
      title: "DOC Tracking",
      path: "/doc-tracking",
      icon: (
        <FaClipboardList className="text-emerald-600 text-2xl sm:text-3xl" />
      ),
      bg: "bg-emerald-50 hover:bg-emerald-100 border-emerald-200",
      textColor: "text-emerald-900",
    },
    {
      title: "Vet Visits",
      path: "/visits",
      icon: <FaStethoscope className="text-indigo-600 text-2xl sm:text-3xl" />,
      bg: "bg-indigo-50 hover:bg-indigo-100 border-indigo-200",
      textColor: "text-indigo-900",
    },
    {
      title: "Reports",
      path: "/reports",
      icon: <FaChartBar className="text-amber-600 text-2xl sm:text-3xl" />,
      bg: "bg-amber-50 hover:bg-amber-100 border-amber-200",
      textColor: "text-amber-900",
    },
    {
      title: "Farmers",
      path: "/farmers",
      icon: <FaUserFriends className="text-teal-600 text-2xl sm:text-3xl" />,
      bg: "bg-teal-50 hover:bg-teal-100 border-teal-200",
      textColor: "text-teal-900",
    },
    {
      title: "Agents",
      path: "/agents",
      icon: <FaUsers className="text-purple-600 text-2xl sm:text-3xl" />,
      bg: "bg-purple-50 hover:bg-purple-100 border-purple-200",
      textColor: "text-purple-900",
    },
  ];

  return (
    <div className="p-3 sm:p-6 max-w-7xl mx-auto space-y-5">
      {/* 5 CARDS (2 COLUMNS EVEN ON MOBILE) */}
      <div className="grid grid-cols-2 gap-3 sm:gap-5">
        {navCards.map((card) => (
          <Link
            key={card.path}
            to={card.path}
            className={`border rounded-2xl p-3 sm:p-5 shadow-sm transition transform hover:-translate-y-0.5 flex flex-col sm:flex-row items-center sm:items-center gap-2 sm:gap-4 text-center sm:text-left ${card.bg}`}
          >
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
          </Link>
        ))}
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
    </div>
  );
}