import React, { useEffect, useState } from "react";
import { Link, useLocation, useNavigate } from "react-router-dom";
import { FaBell, FaBars, FaTimes, FaHome, } from "react-icons/fa";
import { supabase } from "../../lib/supabase";

function getAgeDays(hatchDateStr?: string, refDateStr?: string): number {
  if (!hatchDateStr) return 0;
  const hatchDate = new Date(hatchDateStr);
  const refDate = refDateStr ? new Date(refDateStr) : new Date();
  const diffTime = refDate.getTime() - hatchDate.getTime();
  return Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
}

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

export default function Navbar({
  sidebarOpen,
  setSidebarOpen,
}: {
  role?: string;
  userName?: string;
  sidebarOpen: boolean;
  setSidebarOpen: React.Dispatch<React.SetStateAction<boolean>>;
}) {
  const location = useLocation();
  const navigate = useNavigate();

  const [isModalOpen, setIsModalOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [runningVisits, setRunningVisits] = useState<any[]>([]);
  const [, setAgentsMap] = useState<Record<string, any>>({});
  const [vaccinesList, setVaccinesList] = useState<any[]>([]);
  const [lowWeightList, setLowWeightList] = useState<any[]>([]);
  const [lowEggList, setLowEggList] = useState<any[]>([]);

  const pageTitle =
    location.pathname === "/"
      ? "Dashboard"
      : location.pathname
          .replace("/", "")
          .replaceAll("-", " ")
          .replace(/\b\w/g, (c) => c.toUpperCase());

  useEffect(() => {
    loadNotificationsData();
  }, []);

  async function loadNotificationsData() {
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

      const [
        { data: vData },
        { data: agData },
        { data: frmData },
        { data: vacData },
      ] = await Promise.all([
        supabase
          .from("vet_visits")
          .select("*")
          .in("status", ["Activated", "Running"]),
        supabase.from("agents").select("*"),
        supabase.from("farmers").select("*"),
        supabase.from("vaccines").select("*"),
      ]);

      const agDict: Record<string, any> = {};
      (agData || []).forEach((a) => (agDict[String(a.id)] = a));
      setAgentsMap(agDict);
      setVaccinesList(vacData || []);

      const frmDict: Record<string, any> = {};
      (frmData || []).forEach((f) => (frmDict[String(f.id)] = f));

      let visits = (vData || []).map((v) => ({
        ...v,
        agents: agDict[String(v.agent_id)] || {
          agent_name: "Unknown Agent",
          upazila: "",
        },
        farmers: frmDict[String(v.farmer_id)] || {
          farmer_name: "Unknown Farmer",
        },
      }));

      // Filter by assigned upazila if restricted
      if (!isFullAccess && uUpazilas.length > 0) {
        visits = visits.filter((v) => uUpazilas.includes(v.agents?.upazila));
      }

      setRunningVisits(visits);

      // ---- Below-standard lists (based on each batch's LAST visit report) ----
      const visitIds = visits.map((v) => String(v.id));
      if (visitIds.length > 0) {
        const [
          { data: bRep },
          { data: lRep },
          { data: bStd },
          { data: lStd },
        ] = await Promise.all([
          supabase
            .from("before_sale_reports")
            .select(
              "vet_visit_id, visit_date, weight_gm_b, feed_gm_b, feed_t_kg"
            )
            .in("vet_visit_id", visitIds)
            .order("visit_date", { ascending: false }),
          supabase
            .from("layer_reports")
            .select(
              "vet_visit_id, visit_date, egg_production, feed_intake_g, feed_total_kg"
            )
            .in("vet_visit_id", visitIds)
            .order("visit_date", { ascending: false }),
          supabase.from("broiler_standards").select("*"),
          supabase.from("layer_standards").select("*"),
        ]);

        const bStdMap: Record<number, any> = {};
        (bStd || []).forEach((s: any) => (bStdMap[Number(s.age_days)] = s));

        const lStdMap: Record<string, any> = {};
        (lStd || []).forEach((s: any) => {
          const breedKey = String(s.breed || "shaver star cross 579 brown")
            .trim()
            .toLowerCase();
          lStdMap[`${breedKey}_${Number(s.age_weeks)}`] = s;
          if (!lStdMap[`any_${Number(s.age_weeks)}`]) {
            lStdMap[`any_${Number(s.age_weeks)}`] = s;
          }
        });

        const lastB: Record<string, any> = {};
        (bRep || []).forEach((r: any) => {
          const k = String(r.vet_visit_id);
          if (!lastB[k]) lastB[k] = r; // already sorted desc
        });
        const lastL: Record<string, any> = {};
        (lRep || []).forEach((r: any) => {
          const k = String(r.vet_visit_id);
          if (!lastL[k]) lastL[k] = r;
        });

        const lowW: any[] = [];
        const lowE: any[] = [];
        visits.forEach((v) => {
          const cat = (v.category || "Broiler").toLowerCase();
          if (cat === "broiler") {
            const rep = lastB[String(v.id)];
            if (!rep) return;
            const ageAtVisit = getAgeDays(v.hatch_date, rep.visit_date);
            const std = bStdMap[ageAtVisit];
            const stdWt = Number(std?.weight || 0);
            const actual = Number(rep.weight_gm_b || 0);

            const tBirds = Number(v.t_birds || 0);
            const actualFeed =
              Number(rep.feed_gm_b || 0) ||
              (tBirds > 0 && Number(rep.feed_t_kg || 0) > 0
                ? Number(
                    ((Number(rep.feed_t_kg) * 1000) / tBirds).toFixed(1)
                  )
                : 0);
            const stdFeed = Number(std?.feed_intake || 0);

            if (stdWt > 0 && actual > 0 && actual < stdWt) {
              lowW.push({
                ...v,
                lastVisitDate: rep.visit_date,
                ageAtVisit,
                actual,
                stdWt,
                lessBy: stdWt - actual,
                actualFeed,
                stdFeed,
              });
            }
          } else if (cat === "layer") {
            const rep = lastL[String(v.id)];
            if (!rep) return;
            const ageDays = getAgeDays(v.hatch_date, rep.visit_date);
            const wks = Math.min(100, Math.max(1, Math.floor(ageDays / 7)));
            const isBovans = String(v.breed || "")
              .toLowerCase()
              .includes("bovans");
            const key = isBovans
              ? "bovans white"
              : "shaver star cross 579 brown";
            const std = lStdMap[`${key}_${wks}`] || lStdMap[`any_${wks}`];
            const stdEgg = Number(std?.egg_production || 0);
            const actual = Number(rep.egg_production || 0);

            const tBirds = Number(v.t_birds || 0);
            const actualFeed =
              Number(rep.feed_intake_g || 0) ||
              (tBirds > 0 && Number(rep.feed_total_kg || 0) > 0
                ? Number(
                    ((Number(rep.feed_total_kg) * 1000) / tBirds).toFixed(1)
                  )
                : 0);
            const stdFeed = Number(std?.feed_bird_day || 0);

            if (stdEgg > 0 && actual < stdEgg) {
              lowE.push({
                ...v,
                lastVisitDate: rep.visit_date,
                ageWeeks: wks,
                actual,
                stdEgg,
                lessBy: stdEgg - actual,
                actualFeed,
                stdFeed,
              });
            }
          }
        });
        setLowWeightList(lowW);
        setLowEggList(lowE);
      } else {
        setLowWeightList([]);
        setLowEggList([]);
      }
    } catch (err) {
      console.error("Failed to load notifications:", err);
    } finally {
      setLoading(false);
    }
  }

  // Check Broiler/Sonali Sell Alerts (Age >= 28 days)
  const broilerSellAlerts = runningVisits.filter((v) => {
    const cat = (v.category || "Broiler").toLowerCase();
    if (cat !== "broiler" && cat !== "sonali") return false;
    const age = getAgeDays(v.hatch_date);
    return age >= 28;
  });

  // Check Layer Vaccine Alerts (Within next 3 days)
  const layerVaccineAlerts: any[] = [];
  runningVisits.forEach((v) => {
    const cat = (v.category || "").toLowerCase();
    if (cat !== "layer") return;
    const ageDays = getAgeDays(v.hatch_date);
    const lastGiven = String(v.last_vaccine_name || "").trim().toLowerCase();

    const matchedVac = vaccinesList.filter(
      (vac) =>
        vac.category?.toLowerCase() === "layer" &&
        Number(vac.age_days) >= ageDays &&
        Number(vac.age_days) <= ageDays + 3 &&
        !lastGiven.includes(vac.vaccine_name?.trim().toLowerCase())
    );

    matchedVac.forEach((vac) => {
      layerVaccineAlerts.push({
        ...v,
        vaccineDue: vac,
        daysLeft: Number(vac.age_days) - ageDays,
      });
    });
  });

  const totalAlertsCount =
    broilerSellAlerts.length +
    layerVaccineAlerts.length +
    lowWeightList.length +
    lowEggList.length;

  function groupByAgent(list: any[]) {
    const out: Record<string, any[]> = {};
    list.forEach((v) => {
      const n = v.agents?.agent_name || "Unassigned Agent";
      if (!out[n]) out[n] = [];
      out[n].push(v);
    });
    return out;
  }
  const lowWeightByAgent = groupByAgent(lowWeightList);
  const lowEggByAgent = groupByAgent(lowEggList);

  // Group Broiler Sell Alerts by Agent
  const broilerByAgent: Record<string, any[]> = {};
  broilerSellAlerts.forEach((v) => {
    const agName = v.agents?.agent_name || "Unassigned Agent";
    if (!broilerByAgent[agName]) broilerByAgent[agName] = [];
    broilerByAgent[agName].push(v);
  });

  // Group Layer Vaccine Alerts by Agent
  const layerByAgent: Record<string, any[]> = {};
  layerVaccineAlerts.forEach((v) => {
    const agName = v.agents?.agent_name || "Unassigned Agent";
    if (!layerByAgent[agName]) layerByAgent[agName] = [];
    layerByAgent[agName].push(v);
  });

  return (
    <>
      <header className="sticky top-0 z-40 h-14 sm:h-16 bg-white border-b border-slate-200 shadow-sm flex items-center justify-between px-4 sm:px-6">
        {/* Left: Menu Toggle Icon + Logo + Page Title */}
        <div className="flex items-center gap-3 min-w-0">
          <button
            onClick={() => setSidebarOpen(!sidebarOpen)}
            className="w-10 h-10 rounded-xl bg-slate-100 hover:bg-emerald-50 text-slate-700 hover:text-emerald-600 flex items-center justify-center transition active:scale-95 shrink-0"
            aria-label="Toggle Menu"
          >
            <FaBars size={18} />
          </button>

          <div className="flex items-center gap-3">
            <img
              src="/logo smart team.png"
              alt="FAMS Logo"
              className="w-12 h-12 object-contain"
            />
          </div>

          <h2 className="text-base sm:text-xl font-bold text-slate-800 truncate">
            {pageTitle}
          </h2>

          <Link
              to="/"
              className="w-9 h-9 rounded-xl bg-slate-100 hover:bg-emerald-50 text-slate-700 hover:text-emerald-600 flex items-center justify-center transition active:scale-95 shrink-0"
              title="Home"
            >
              <FaHome size={16} />
            </Link>
            
        </div>

        {/* Right: Notification Icon with Badge */}
        <div className="flex items-center">
          <button
            onClick={() => {
              loadNotificationsData();
              setIsModalOpen(true);
            }}
            className="relative w-10 h-10 rounded-xl hover:bg-slate-100 text-slate-600 hover:text-emerald-600 flex items-center justify-center transition"
            aria-label="Notifications"
          >
            <FaBell size={19} />
            {totalAlertsCount > 0 && (
              <span className="absolute top-2 right-2.5 h-2.5 w-2.5 rounded-full bg-red-500 ring-2 ring-white animate-pulse" />
            )}
          </button>
        </div>
      </header>

      {/* ==================================================
          NOTIFICATION MODAL POPUP
      ================================================== */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-3">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden border border-slate-200">
            <div className="flex items-center justify-between px-5 py-4 border-b border-slate-200 bg-slate-100">
              <div>
                <h3 className="text-base sm:text-lg font-bold text-slate-800 flex items-center gap-2">
                  <FaBell className="text-emerald-600" />
                  Field Alerts & Notifications
                </h3>
                <p className="text-xs text-slate-500">
                  Total Active Alerts: <strong>{totalAlertsCount}</strong>
                </p>
              </div>
              <button
                onClick={() => setIsModalOpen(false)}
                className="w-8 h-8 rounded-lg hover:bg-slate-200 flex items-center justify-center text-slate-500"
              >
                <FaTimes size={16} />
              </button>
            </div>

            <div className="p-4 sm:p-6 space-y-6 overflow-y-auto flex-1 text-xs sm:text-sm">
              {loading ? (
                <div className="py-12 text-center text-slate-500">
                  Checking active batch alerts...
                </div>
              ) : totalAlertsCount === 0 ? (
                <div className="py-12 text-center text-slate-400 font-medium">
                  No pending sale or vaccination alerts at this moment.
                </div>
              ) : (
                <>
                  {/* SECTION 1: BROILER & SONALI SALE TIME ALERTS */}
                  {Object.keys(broilerByAgent).length > 0 && (
                    <div className="space-y-3">
                      <h4 className="font-bold text-red-700 text-sm sm:text-base uppercase tracking-wider border-b pb-1">
                        Ready for Sale
                      </h4>

                      {Object.entries(broilerByAgent).map(([agentName, list]) => (
                        <div
                          key={agentName}
                          className="bg-slate-50 border border-slate-200 rounded-xl overflow-hidden shadow-sm"
                        >
                          <div className="bg-slate-100 px-4 py-2 font-bold text-slate-800 border-b border-slate-200">
                            Agent:{" "}
                            <span className="text-emerald-700">{agentName}</span>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse">
                              <thead>
                                <tr className="bg-slate-50/70 border-b border-slate-200 text-slate-600">
                                  <th className="px-4 py-2">Farmer Name</th>
                                  <th className="px-4 py-2">Category</th>
                                  <th className="px-4 py-2">Hatch Date</th>
                                  <th className="px-4 py-2 font-bold text-red-600">
                                    Age
                                  </th>
                                  <th className="px-4 py-2 text-right">
                                    Action
                                  </th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                {list.map((v) => {
                                  const age = getAgeDays(v.hatch_date);
                                  return (
                                    <tr
                                      key={v.id}
                                      className="hover:bg-slate-100/50"
                                    >
                                      <td className="px-4 py-2.5 font-bold text-slate-800">
                                        {v.batch_label ||
                                          v.farmers?.farmer_name}
                                      </td>
                                      <td className="px-4 py-2.5 font-semibold text-slate-700">
                                        {v.category || "Broiler"}
                                      </td>
                                      <td className="px-4 py-2.5 text-slate-600">
                                        {v.hatch_date || "-"}
                                      </td>
                                      <td className="px-4 py-2.5 font-bold text-red-600">
                                        {age} Days
                                      </td>
                                      <td className="px-4 py-2.5 text-right">
                                        <button
                                          type="button"
                                          onClick={() => {
                                            setIsModalOpen(false);
                                            navigate(
                                              `/visits?focus=${v.id}&action=after_sale`
                                            );
                                          }}
                                          className="bg-red-600 hover:bg-red-700 text-white px-3 py-1 rounded-lg text-xs font-bold shadow-sm transition"
                                        >
                                          Sold
                                        </button>
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                  {/* SECTION 2: LAYER VACCINE ALERTS (WITHIN 3 DAYS) */}
                  {Object.keys(layerByAgent).length > 0 && (
                    <div className="space-y-3 pt-2">
                      <h4 className="font-bold text-indigo-700 text-sm sm:text-base uppercase tracking-wider border-b pb-1">
                        Layer Upcoming Vaccines (in 3 days)
                      </h4>

                      {Object.entries(layerByAgent).map(([agentName, list]) => (
                        <div
                          key={agentName}
                          className="bg-indigo-50/30 border border-indigo-200 rounded-xl overflow-hidden shadow-sm"
                        >
                          <div className="bg-indigo-50 px-4 py-2 font-bold text-indigo-900 border-b border-indigo-200">
                            Agent:{" "}
                            <span className="text-emerald-700">{agentName}</span>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse">
                              <thead>
                                <tr className="bg-indigo-50/50 border-b border-indigo-200 text-indigo-900">
                                  <th className="px-4 py-2">Farmer Name</th>
                                  <th className="px-4 py-2">Current Age</th>
                                  <th className="px-4 py-2">Vaccine Name</th>
                                  <th className="px-4 py-2">Route</th>
                                  <th className="px-4 py-2 text-right">
                                    Due Schedule
                                  </th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-indigo-100">
                                {list.map((v, i) => {
                                  const ageDays = getAgeDays(v.hatch_date);
                                  return (
                                    <tr
                                      key={i}
                                      className="hover:bg-indigo-50/40"
                                    >
                                      <td className="px-4 py-2.5 font-bold text-slate-800">
                                        {v.batch_label ||
                                          v.farmers?.farmer_name}
                                      </td>
                                      <td className="px-4 py-2.5 font-semibold text-slate-700">
                                        {ageDays} Days
                                      </td>
                                      <td className="px-4 py-2.5 font-bold text-indigo-800">
                                        {v.vaccineDue.vaccine_name}
                                      </td>
                                      <td className="px-4 py-2.5 text-slate-600">
                                        {v.vaccineDue.route}
                                      </td>
                                      <td className="px-4 py-2.5 text-right font-bold text-amber-700">
                                        Day {v.vaccineDue.age_days} (
                                        {v.daysLeft === 0
                                          ? "Today"
                                          : `in ${v.daysLeft} days`}
                                        )
                                      </td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}

                  {/* SECTION 3: BROILER BELOW STANDARD WEIGHT (LAST VISIT) */}
                  {Object.keys(lowWeightByAgent).length > 0 && (
                    <div className="space-y-3 pt-2">
                      <h4 className="font-bold text-orange-700 text-sm sm:text-base uppercase tracking-wider border-b pb-1">
                        Broiler — Below Standard Weight
                      </h4>
                      {Object.entries(lowWeightByAgent).map(
                        ([agentName, list]) => (
                          <div
                            key={agentName}
                            className="bg-orange-50/30 border border-orange-200 rounded-xl overflow-hidden shadow-sm"
                          >
                            <div className="bg-orange-50 px-4 py-2 font-bold text-orange-900 border-b border-orange-200">
                              Agent:{" "}
                              <span className="text-emerald-700">
                                {agentName}
                              </span>
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-left border-collapse">
                                <thead>
                                  <tr className="bg-orange-50/50 border-b border-orange-200 text-orange-900">
                                    <th className="px-3 py-2">Farmer</th>
                                    <th className="px-3 py-2">Last Visit</th>
                                    <th className="px-3 py-2">Age</th>
                                    <th className="px-3 py-2 text-right">
                                      Feed Intake (g/b)
                                    </th>
                                    <th className="px-3 py-2 text-right">
                                      Std Feed (g/b)
                                    </th>
                                    <th className="px-3 py-2 text-right">
                                      Weight
                                    </th>
                                    <th className="px-3 py-2 text-right">
                                      Std Wt
                                    </th>
                                    <th className="px-3 py-2 text-right font-bold text-red-600">
                                      Less than Std
                                    </th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-orange-100">
                                  {list.map((v) => (
                                    <tr
                                      key={v.id}
                                      className="hover:bg-orange-50/40"
                                    >
                                      <td className="px-3 py-2 font-bold text-slate-800">
                                        {v.batch_label ||
                                          v.farmers?.farmer_name}
                                      </td>
                                      <td className="px-3 py-2 text-slate-600 whitespace-nowrap">
                                        {v.lastVisitDate || "-"}
                                      </td>
                                      <td className="px-3 py-2 text-slate-700">
                                        {v.ageAtVisit}d
                                      </td>
                                      <td className="px-3 py-2 text-right font-semibold text-slate-800">
                                        {v.actualFeed > 0
                                          ? `${v.actualFeed} g`
                                          : "-"}
                                      </td>
                                      <td className="px-3 py-2 text-right text-slate-600">
                                        {v.stdFeed > 0 ? `${v.stdFeed} g` : "-"}
                                      </td>
                                      <td className="px-3 py-2 text-right font-semibold">
                                        {v.actual} g
                                      </td>
                                      <td className="px-3 py-2 text-right text-slate-600">
                                        {v.stdWt} g
                                      </td>
                                      <td className="px-3 py-2 text-right font-bold text-red-600">
                                        -{v.lessBy.toFixed(0)} g
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )
                      )}
                    </div>
                  )}

                  {/* SECTION 4: LAYER BELOW STANDARD EGG PRODUCTION (LAST VISIT) */}
                  {Object.keys(lowEggByAgent).length > 0 && (
                    <div className="space-y-3 pt-2">
                      <h4 className="font-bold text-yellow-700 text-sm sm:text-base uppercase tracking-wider border-b pb-1">
                        Layer — Below Standard Egg Production
                      </h4>
                      {Object.entries(lowEggByAgent).map(
                        ([agentName, list]) => (
                          <div
                            key={agentName}
                            className="bg-yellow-50/30 border border-yellow-200 rounded-xl overflow-hidden shadow-sm"
                          >
                            <div className="bg-yellow-50 px-4 py-2 font-bold text-yellow-900 border-b border-yellow-200">
                              Agent:{" "}
                              <span className="text-emerald-700">
                                {agentName}
                              </span>
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-left border-collapse">
                                <thead>
                                  <tr className="bg-yellow-50/50 border-b border-yellow-200 text-yellow-900">
                                    <th className="px-3 py-2">Farmer</th>
                                    <th className="px-3 py-2">Last Visit</th>
                                    <th className="px-3 py-2">Age</th>
                                    <th className="px-3 py-2 text-right">
                                      Feed Intake (g/b)
                                    </th>
                                    <th className="px-3 py-2 text-right">
                                      Std Feed (g/b)
                                    </th>
                                    <th className="px-3 py-2 text-right">
                                      Egg %
                                    </th>
                                    <th className="px-3 py-2 text-right">
                                      Std %
                                    </th>
                                    <th className="px-3 py-2 text-right font-bold text-red-600">
                                      Less than Std
                                    </th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-yellow-100">
                                  {list.map((v) => (
                                    <tr
                                      key={v.id}
                                      className="hover:bg-yellow-50/40"
                                    >
                                      <td className="px-3 py-2 font-bold text-slate-800">
                                        {v.batch_label ||
                                          v.farmers?.farmer_name}
                                      </td>
                                      <td className="px-3 py-2 text-slate-600 whitespace-nowrap">
                                        {v.lastVisitDate || "-"}
                                      </td>
                                      <td className="px-3 py-2 text-slate-700">
                                        {v.ageWeeks} wks
                                      </td>
                                      <td className="px-3 py-2 text-right font-semibold text-slate-800">
                                        {v.actualFeed > 0
                                          ? `${v.actualFeed} g`
                                          : "-"}
                                      </td>
                                      <td className="px-3 py-2 text-right text-slate-600">
                                        {v.stdFeed > 0 ? `${v.stdFeed} g` : "-"}
                                      </td>
                                      <td className="px-3 py-2 text-right font-semibold">
                                        {v.actual}%
                                      </td>
                                      <td className="px-3 py-2 text-right text-slate-600">
                                        {v.stdEgg}%
                                      </td>
                                      <td className="px-3 py-2 text-right font-bold text-red-600">
                                        -{v.lessBy.toFixed(1)}%
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )
                      )}
                    </div>
                  )}
                </>
              )}
            </div>

            <div className="px-5 py-3 border-t border-slate-200 bg-slate-50 flex justify-end">
              <button
                onClick={() => setIsModalOpen(false)}
                className="px-5 py-2 rounded-xl bg-slate-800 hover:bg-slate-900 text-white text-xs sm:text-sm font-semibold transition"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}