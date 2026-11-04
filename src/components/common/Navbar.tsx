import React, { useEffect, useState } from "react";
import { useLocation, useNavigate } from "react-router-dom";
import { FaBell, FaBars, FaTimes } from "react-icons/fa";
import { supabase } from "../../lib/supabase";

function getAgeDays(hatchDateStr?: string): number {
  if (!hatchDateStr) return 0;
  const hatchDate = new Date(hatchDateStr);
  const diffTime = new Date().getTime() - hatchDate.getTime();
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
        agents: agDict[String(v.agent_id)] || { agent_name: "Unknown Agent", upazila: "" },
        farmers: frmDict[String(v.farmer_id)] || { farmer_name: "Unknown Farmer" },
      }));

      // Filter by assigned upazila if restricted
      if (!isFullAccess && uUpazilas.length > 0) {
        visits = visits.filter((v) => uUpazilas.includes(v.agents?.upazila));
      }

      setRunningVisits(visits);
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

  const totalAlertsCount = broilerSellAlerts.length + layerVaccineAlerts.length;

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
                            Agent: <span className="text-emerald-700">{agentName}</span>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse">
                              <thead>
                                <tr className="bg-slate-50/70 border-b border-slate-200 text-slate-600">
                                  <th className="px-4 py-2">Farmer Name</th>
                                  <th className="px-4 py-2">Category</th>
                                  <th className="px-4 py-2">Hatch Date</th>
                                  <th className="px-4 py-2 font-bold text-red-600">Age</th>
                                  <th className="px-4 py-2 text-right">Action</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100">
                                {list.map((v) => {
                                  const age = getAgeDays(v.hatch_date);
                                  return (
                                    <tr key={v.id} className="hover:bg-slate-100/50">
                                      <td className="px-4 py-2.5 font-bold text-slate-800">
                                        {v.batch_label || v.farmers?.farmer_name}
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
                                            navigate("/visits");
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
                        Layer Upcoming Vaccines
                      </h4>

                      {Object.entries(layerByAgent).map(([agentName, list]) => (
                        <div
                          key={agentName}
                          className="bg-indigo-50/30 border border-indigo-200 rounded-xl overflow-hidden shadow-sm"
                        >
                          <div className="bg-indigo-50 px-4 py-2 font-bold text-indigo-900 border-b border-indigo-200">
                            Agent: <span className="text-emerald-700">{agentName}</span>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse">
                              <thead>
                                <tr className="bg-indigo-50/50 border-b border-indigo-200 text-indigo-900">
                                  <th className="px-4 py-2">Farmer Name</th>
                                  <th className="px-4 py-2">Current Age</th>
                                  <th className="px-4 py-2">Vaccine Name</th>
                                  <th className="px-4 py-2">Route</th>
                                  <th className="px-4 py-2 text-right">Due Schedule</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-indigo-100">
                                {list.map((v, i) => {
                                  const ageDays = getAgeDays(v.hatch_date);
                                  return (
                                    <tr key={i} className="hover:bg-indigo-50/40">
                                      <td className="px-4 py-2.5 font-bold text-slate-800">
                                        {v.batch_label || v.farmers?.farmer_name}
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
                                        Day {v.vaccineDue.age_days} ({v.daysLeft === 0 ? "Today" : `in ${v.daysLeft} days`})
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