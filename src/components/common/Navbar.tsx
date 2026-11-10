import React, { useEffect, useRef, useState } from "react";
import { Link, useLocation, useNavigate } from "react-router-dom";
import { FaBell, FaBars, FaTimes, FaHome } from "react-icons/fa";
import { LocalNotifications } from "@capacitor/local-notifications";
import { Capacitor } from "@capacitor/core";
import { supabase } from "../../lib/supabase";
import {
  initNativePermissions,
  triggerPhoneNotification,
} from "../../utils/nativeAppHelper";

function getAgeDays(hatchDateStr?: string, refDateStr?: string): number {
  if (!hatchDateStr) return 0;
  const hatchDate = new Date(hatchDateStr);
  const refDate = refDateStr ? new Date(refDateStr) : new Date();
  const diffTime = refDate.getTime() - hatchDate.getTime();
  return Math.max(0, Math.floor(diffTime / (1000 * 60 * 60 * 24)));
}

function parseArray(field: any): string[] {
  if (!field) return [];
  if (Array.isArray(field)) return field.map(String).filter(Boolean);
  if (typeof field === "string") {
    try {
      const parsed = JSON.parse(field);
      if (Array.isArray(parsed)) return parsed.map(String).filter(Boolean);
    } catch {
      return field
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
    }
  }
  return [];
}

function timeAgoShort(iso?: string): string {
  if (!iso) return "";
  const diff = Date.now() - new Date(iso).getTime();
  const min = Math.floor(diff / 60000);
  if (min < 1) return "just now";
  if (min < 60) return `${min}m ago`;
  const hr = Math.floor(min / 60);
  if (hr < 24) return `${hr}h ago`;
  return `${Math.floor(hr / 24)}d ago`;
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
  const [latestTrackings, setLatestTrackings] = useState<any[]>([]);

  const farmersRef = useRef<Record<string, any>>({});
  const accessRef = useRef<{
    isFullAccess: boolean;
    upazilas: string[];
    districts: string[];
  }>({
    isFullAccess: false,
    upazilas: [],
    districts: [],
  });

  const pageTitle =
    location.pathname === "/"
      ? "Dashboard"
      : location.pathname
          .replace("/", "")
          .replaceAll("-", " ")
          .replace(/\b\w/g, (c) => c.toUpperCase());

  function isFarmerAllowed(farmer: any): boolean {
    if (!farmer) return false;
    const { isFullAccess, upazilas, districts } = accessRef.current;
    if (isFullAccess) return true;
    const matchUpa =
      upazilas.length === 0 || upazilas.includes(farmer.upazila);
    const matchDist =
      districts.length === 0 || districts.includes(farmer.district);
    return matchUpa && matchDist;
  }

  async function refreshLatestTrackings() {
    const { data: trkData } = await supabase
      .from("farmer_trackings")
      .select("*")
      .order("created_at", { ascending: false })
      .limit(40);

    const list: any[] = [];
    for (const t of trkData || []) {
      const f = farmersRef.current[String(t.farmer_id)];
      if (isFarmerAllowed(f)) {
        list.push({
          ...t,
          farmer_name: f?.farmer_name || "Unknown Farmer",
          upazila: f?.upazila || "",
        });
        if (list.length >= 5) break;
      }
    }
    setLatestTrackings(list);
  }

  useEffect(() => {
    initNativePermissions();
    loadNotificationsData();

    // ১. ফোনের নোটিফিকেশন প্যানেলে ক্লিক করলে নির্দিষ্ট পেজে ও স্ক্রিনে ফোকাস করা
    let notifSub: any = null;
    if (Capacitor.isNativePlatform()) {
      LocalNotifications.addListener(
        "localNotificationActionPerformed",
        (action) => {
          const targetRoute = action.notification.extra?.route;
          if (targetRoute) {
            setIsModalOpen(false);
            navigate(targetRoute);
          }
        }
      ).then((sub) => {
        notifSub = sub;
      });
    }

    // ২. Live Farmer Tracking Realtime Listener + Phone Notification
    const channel = supabase
      .channel("navbar_live_farmer_trackings")
      .on(
        "postgres_changes",
        { event: "*", schema: "public", table: "farmer_trackings" },
        async (payload: any) => {
          if (payload.eventType === "INSERT" && payload.new) {
            const newRow: any = payload.new;
            let farmer = farmersRef.current[String(newRow.farmer_id)];
            if (!farmer && newRow.farmer_id) {
              const { data: fData } = await supabase
                .from("farmers")
                .select("*")
                .eq("id", newRow.farmer_id)
                .single();
              if (fData) {
                farmersRef.current[String(fData.id)] = fData;
                farmer = fData;
              }
            }
            if (isFarmerAllowed(farmer)) {
              const fName = farmer?.farmer_name || "Unknown Farmer";
              const enriched = {
                ...newRow,
                farmer_name: fName,
                upazila: farmer?.upazila || "",
              };
              setLatestTrackings((prev) => {
                const withoutDup = prev.filter(
                  (item) => String(item.id) !== String(enriched.id)
                );
                return [enriched, ...withoutDup].slice(0, 5);
              });

              // ফোনের নোটিফিকেশনে সাথে সাথে পাঠানো
              triggerPhoneNotification(
                `Farmer Tracking: ${fName}`,
                `${newRow.user_name || "Officer"}: ${newRow.note}`,
                `/farmer-tracking?focus=${newRow.farmer_id}&note=${newRow.id}`
              );
            }
          } else {
            refreshLatestTrackings();
          }
        }
      )
      .subscribe();

    return () => {
      supabase.removeChannel(channel);
      if (notifSub) notifSub.remove();
    };
  }, []);

  async function loadNotificationsData() {
    setLoading(true);
    try {
      const {
        data: { user },
      } = await supabase.auth.getUser();

      let uUpazilas: string[] = [];
      let uDistricts: string[] = [];
      let role = "officer";

      if (user) {
        const { data: profile } = await supabase
          .from("users")
          .select("role, upazilas, districts")
          .eq("id", user.id)
          .single();

        if (profile) {
          role = profile.role || "officer";
          uUpazilas = parseArray(profile.upazilas);
          uDistricts = parseArray(profile.districts);
        }
      }

      const isFullAccess = role === "super_admin" || role === "manager";
      accessRef.current = {
        isFullAccess,
        upazilas: uUpazilas,
        districts: uDistricts,
      };

      const [
        { data: vData },
        { data: agData },
        { data: frmData },
        { data: vacData },
        { data: trkData },
      ] = await Promise.all([
        supabase
          .from("vet_visits")
          .select("*")
          .in("status", ["Activated", "Running"]),
        supabase.from("agents").select("*"),
        supabase.from("farmers").select("*"),
        supabase.from("vaccines").select("*"),
        supabase
          .from("farmer_trackings")
          .select("*")
          .order("created_at", { ascending: false })
          .limit(40),
      ]);

      const agDict: Record<string, any> = {};
      (agData || []).forEach((a: any) => (agDict[String(a.id)] = a));
      setAgentsMap(agDict);
      setVaccinesList(vacData || []);

      const frmDict: Record<string, any> = {};
      (frmData || []).forEach((f: any) => (frmDict[String(f.id)] = f));
      farmersRef.current = frmDict;

      const latest5: any[] = [];
      for (const t of trkData || []) {
        const f = frmDict[String(t.farmer_id)];
        if (isFarmerAllowed(f)) {
          latest5.push({
            ...t,
            farmer_name: f?.farmer_name || "Unknown Farmer",
            upazila: f?.upazila || "",
          });
          if (latest5.length >= 5) break;
        }
      }
      setLatestTrackings(latest5);

      let visits = (vData || []).map((v: any) => ({
        ...v,
        agents: agDict[String(v.agent_id)] || {
          agent_name: "Unknown Agent",
          upazila: "",
        },
        farmers: frmDict[String(v.farmer_id)] || {
          farmer_name: "Unknown Farmer",
        },
      }));

      if (!isFullAccess && uUpazilas.length > 0) {
        visits = visits.filter((v: any) =>
          uUpazilas.includes(v.agents?.upazila)
        );
      }

      setRunningVisits(visits);

      // ৩. Broiler/Sonali Time to Sale (Age >= 28) ফোনের নোটিফিকেশনে পাঠানো (একবার করে প্রতি দিন/ব্যাচে)
      const todayKey = new Date().toISOString().slice(0, 10);
      visits.forEach((v: any) => {
        const cat = (v.category || "Broiler").toLowerCase();
        if (cat !== "broiler" && cat !== "sonali") return;
        const age = getAgeDays(v.hatch_date);
        if (age >= 28) {
          const storageKey = `notified_sale_${v.id}_${todayKey}`;
          if (!localStorage.getItem(storageKey)) {
            localStorage.setItem(storageKey, "1");
            const fName =
              v.batch_label || v.farmers?.farmer_name || "Unknown Farmer";
            triggerPhoneNotification(
              "Broiler Time to sale",
              `Farmer Name: ${fName} | Age: ${age} Days`,
              `/visits?focus=${v.id}&action=after_sale`
            );
          }
        }
      });

      const visitIds = visits.map((v: any) => String(v.id));
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
          if (!lastB[k]) lastB[k] = r;
        });
        const lastL: Record<string, any> = {};
        (lRep || []).forEach((r: any) => {
          const k = String(r.vet_visit_id);
          if (!lastL[k]) lastL[k] = r;
        });

        const lowW: any[] = [];
        const lowE: any[] = [];
        visits.forEach((v: any) => {
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

  const broilerSellAlerts = runningVisits.filter((v: any) => {
    const cat = (v.category || "Broiler").toLowerCase();
    if (cat !== "broiler" && cat !== "sonali") return false;
    const age = getAgeDays(v.hatch_date);
    return age >= 28;
  });

  const layerVaccineAlerts: any[] = [];
  runningVisits.forEach((v: any) => {
    const cat = (v.category || "").toLowerCase();
    if (cat !== "layer") return;
    const ageDays = getAgeDays(v.hatch_date);
    const lastGiven = String(v.last_vaccine_name || "").trim().toLowerCase();

    const matchedVac = vaccinesList.filter(
      (vac: any) =>
        vac.category?.toLowerCase() === "layer" &&
        Number(vac.age_days) >= ageDays &&
        Number(vac.age_days) <= ageDays + 3 &&
        !lastGiven.includes(vac.vaccine_name?.trim().toLowerCase())
    );

    matchedVac.forEach((vac: any) => {
      layerVaccineAlerts.push({
        ...v,
        vaccineDue: vac,
        daysLeft: Number(vac.age_days) - ageDays,
      });
    });
  });

  const totalAlertsCount =
    latestTrackings.length +
    broilerSellAlerts.length +
    layerVaccineAlerts.length +
    lowWeightList.length +
    lowEggList.length;

  function groupByAgent(list: any[]) {
    const out: Record<string, any[]> = {};
    list.forEach((v: any) => {
      const n = v.agents?.agent_name || "Unassigned Agent";
      if (!out[n]) out[n] = [];
      out[n].push(v);
    });
    return out;
  }
  const lowWeightByAgent = groupByAgent(lowWeightList);
  const lowEggByAgent = groupByAgent(lowEggList);
  const broilerByAgent = groupByAgent(broilerSellAlerts);
  const layerByAgent = groupByAgent(layerVaccineAlerts);

  return (
    <>
      <header className="sticky top-0 z-40 h-14 sm:h-16 bg-white border-b border-slate-200 shadow-sm flex items-center justify-between px-4 sm:px-6">
        <div className="flex items-center gap-3 min-w-0">
          <button
            onClick={() => setSidebarOpen(!sidebarOpen)}
            className="w-10 h-10 rounded-xl bg-slate-100 hover:bg-emerald-50 text-slate-700 hover:text-emerald-600 flex items-center justify-center transition active:scale-95 shrink-0"
            aria-label="Toggle Menu"
          >
            <FaBars size={18} />
          </button>

          <div className="flex items-center gap-3 shrink-0">
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

      {isModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-2.5 sm:p-3">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] flex flex-col overflow-hidden border border-slate-200">
            <div className="flex items-center justify-between px-4 py-2.5 border-b border-slate-200 bg-slate-100">
              <div>
                <h3 className="text-sm sm:text-base font-bold text-slate-800 flex items-center gap-2">
                  <FaBell className="text-emerald-600" />
                  Field Alerts & Notifications
                </h3>
                <p className="text-[11px] text-slate-500">
                  Total Active Alerts: <strong>{totalAlertsCount}</strong>
                </p>
              </div>
              <button
                onClick={() => setIsModalOpen(false)}
                className="w-7 h-7 rounded-lg hover:bg-slate-200 flex items-center justify-center text-slate-500"
              >
                <FaTimes size={15} />
              </button>
            </div>

            <div className="p-3 sm:p-4 space-y-3.5 overflow-y-auto flex-1 text-[11px] sm:text-xs">
              {loading && totalAlertsCount === 0 ? (
                <div className="py-10 text-center text-slate-500">
                  Checking active batch alerts...
                </div>
              ) : totalAlertsCount === 0 ? (
                <div className="py-10 text-center text-slate-400 font-medium">
                  No pending sale, vaccination, or tracking alerts at this
                  moment.
                </div>
              ) : (
                <>
                  {/* SECTION 0: LIVE FARMER TRACKING MESSAGES (LATEST 5) */}
                  {latestTrackings.length > 0 && (
                    <div className="space-y-1.5">
                      <div className="flex items-center justify-between border-b border-emerald-200 pb-0.5">
                        <h4 className="font-bold text-emerald-800 text-xs sm:text-sm uppercase tracking-wider flex items-center gap-1.5">
                          <span className="h-2 w-2 rounded-full bg-emerald-500 animate-ping" />
                          Latest Farmer Tracking (Live)
                        </h4>
                        <button
                          type="button"
                          onClick={() => {
                            setIsModalOpen(false);
                            navigate("/farmer-tracking");
                          }}
                          className="text-[11px] font-bold text-emerald-700 hover:underline"
                        >
                          View All →
                        </button>
                      </div>

                      <div className="bg-emerald-50/40 border border-emerald-200 rounded-xl overflow-hidden shadow-xs">
                        <div className="overflow-x-auto">
                          <table className="w-full text-left border-collapse text-[11px] sm:text-xs">
                            <thead>
                              <tr className="bg-emerald-100/70 border-b border-emerald-200 text-emerald-950">
                                <th className="px-2.5 py-1 font-bold whitespace-nowrap">
                                  Farmer Name
                                </th>
                                <th className="px-2.5 py-1 font-bold">
                                  Message
                                </th>
                                <th className="px-2.5 py-1 font-bold text-right whitespace-nowrap">
                                  Message By
                                </th>
                              </tr>
                            </thead>
                            <tbody className="divide-y divide-emerald-100">
                              {latestTrackings.map((t: any) => (
                                <tr
                                  key={t.id}
                                  onClick={() => {
                                    setIsModalOpen(false);
                                    navigate(
                                      `/farmer-tracking?focus=${t.farmer_id}&note=${t.id}`
                                    );
                                  }}
                                  className="hover:bg-emerald-100/60 cursor-pointer leading-tight transition"
                                >
                                  <td className="px-2.5 py-1 font-bold text-slate-800 whitespace-nowrap">
                                    {t.farmer_name}
                                  </td>
                                  <td className="px-2.5 py-1 text-slate-800 font-medium max-w-[220px] sm:max-w-md truncate">
                                    {t.note}
                                  </td>
                                  <td className="px-2.5 py-1 text-right whitespace-nowrap">
                                    <span className="font-bold text-emerald-800">
                                      {t.user_name || "Officer"}
                                    </span>
                                    <span className="ml-1 text-[10px] text-slate-500">
                                      ({timeAgoShort(t.created_at)})
                                    </span>
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  )}

                  {/* SECTION 1: BROILER & SONALI SALE TIME ALERTS */}
                  {Object.keys(broilerByAgent).length > 0 && (
                    <div className="space-y-1.5">
                      <h4 className="font-bold text-red-700 text-xs sm:text-sm uppercase tracking-wider border-b pb-0.5">
                        Ready for Sale
                      </h4>

                      {Object.entries(broilerByAgent).map(
                        ([agentName, list]) => (
                          <div
                            key={agentName}
                            className="bg-slate-50 border border-slate-200 rounded-xl overflow-hidden shadow-xs"
                          >
                            <div className="bg-slate-100 px-3 py-1 font-bold text-slate-800 border-b border-slate-200">
                              Agent:{" "}
                              <span className="text-emerald-700">
                                {agentName}
                              </span>
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-left border-collapse text-[11px] sm:text-xs">
                                <thead>
                                  <tr className="bg-slate-50/70 border-b border-slate-200 text-slate-600">
                                    <th className="px-2.5 py-1">Farmer Name</th>
                                    <th className="px-2.5 py-1">Category</th>
                                    <th className="px-2.5 py-1">Hatch Date</th>
                                    <th className="px-2.5 py-1 font-bold text-red-600">
                                      Age
                                    </th>
                                    <th className="px-2.5 py-1 text-right">
                                      Action
                                    </th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-slate-100">
                                  {list.map((v: any) => {
                                    const age = getAgeDays(v.hatch_date);
                                    return (
                                      <tr
                                        key={v.id}
                                        className="hover:bg-slate-100/50 leading-tight"
                                      >
                                        <td className="px-2.5 py-0.5 font-bold text-slate-800 whitespace-nowrap">
                                          {v.batch_label ||
                                            v.farmers?.farmer_name}
                                        </td>
                                        <td className="px-2.5 py-0.5 font-semibold text-slate-700 whitespace-nowrap">
                                          {v.category || "Broiler"}
                                        </td>
                                        <td className="px-2.5 py-0.5 text-slate-600 whitespace-nowrap">
                                          {v.hatch_date || "-"}
                                        </td>
                                        <td className="px-2.5 py-0.5 font-bold text-red-600 whitespace-nowrap">
                                          {age} Days
                                        </td>
                                        <td className="px-2.5 py-0.5 text-right whitespace-nowrap">
                                          <button
                                            type="button"
                                            onClick={() => {
                                              setIsModalOpen(false);
                                              navigate(
                                                `/visits?focus=${v.id}&action=after_sale`
                                              );
                                            }}
                                            className="bg-red-600 hover:bg-red-700 text-white px-2 py-0.5 rounded text-[10px] font-bold shadow-xs transition"
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
                        )
                      )}
                    </div>
                  )}

                  {/* SECTION 2: LAYER VACCINE ALERTS */}
                  {Object.keys(layerByAgent).length > 0 && (
                    <div className="space-y-1.5 pt-1">
                      <h4 className="font-bold text-indigo-700 text-xs sm:text-sm uppercase tracking-wider border-b pb-0.5">
                        Layer Upcoming Vaccines (in 3 days)
                      </h4>

                      {Object.entries(layerByAgent).map(([agentName, list]) => (
                        <div
                          key={agentName}
                          className="bg-indigo-50/30 border border-indigo-200 rounded-xl overflow-hidden shadow-xs"
                        >
                          <div className="bg-indigo-50 px-3 py-1 font-bold text-indigo-900 border-b border-indigo-200">
                            Agent:{" "}
                            <span className="text-emerald-700">
                              {agentName}
                            </span>
                          </div>
                          <div className="overflow-x-auto">
                            <table className="w-full text-left border-collapse text-[11px] sm:text-xs">
                              <thead>
                                <tr className="bg-indigo-50/50 border-b border-indigo-200 text-indigo-900">
                                  <th className="px-2.5 py-1">Farmer Name</th>
                                  <th className="px-2.5 py-1">Current Age</th>
                                  <th className="px-2.5 py-1">Vaccine Name</th>
                                  <th className="px-2.5 py-1">Route</th>
                                  <th className="px-2.5 py-1 text-right">
                                    Due Schedule
                                  </th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-indigo-100">
                                {list.map((v: any, i: number) => {
                                  const ageDays = getAgeDays(v.hatch_date);
                                  return (
                                    <tr
                                      key={i}
                                      className="hover:bg-indigo-50/40 leading-tight"
                                    >
                                      <td className="px-2.5 py-0.5 font-bold text-slate-800 whitespace-nowrap">
                                        {v.batch_label ||
                                          v.farmers?.farmer_name}
                                      </td>
                                      <td className="px-2.5 py-0.5 font-semibold text-slate-700 whitespace-nowrap">
                                        {ageDays} Days
                                      </td>
                                      <td className="px-2.5 py-0.5 font-bold text-indigo-800 whitespace-nowrap">
                                        {v.vaccineDue.vaccine_name}
                                      </td>
                                      <td className="px-2.5 py-0.5 text-slate-600 whitespace-nowrap">
                                        {v.vaccineDue.route}
                                      </td>
                                      <td className="px-2.5 py-0.5 text-right font-bold text-amber-700 whitespace-nowrap">
                                        Day {v.vaccineDue.age_days} (
                                        {v.daysLeft === 0
                                          ? "Today"
                                          : `in ${v.daysLeft}d`}
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

                  {/* SECTION 3: BROILER BELOW STANDARD WEIGHT */}
                  {Object.keys(lowWeightByAgent).length > 0 && (
                    <div className="space-y-1.5 pt-1">
                      <h4 className="font-bold text-orange-700 text-xs sm:text-sm uppercase tracking-wider border-b pb-0.5">
                        Broiler — Below Standard Weight
                      </h4>
                      {Object.entries(lowWeightByAgent).map(
                        ([agentName, list]) => (
                          <div
                            key={agentName}
                            className="bg-orange-50/30 border border-orange-200 rounded-xl overflow-hidden shadow-xs"
                          >
                            <div className="bg-orange-50 px-3 py-1 font-bold text-orange-900 border-b border-orange-200">
                              Agent:{" "}
                              <span className="text-emerald-700">
                                {agentName}
                              </span>
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-left border-collapse text-[11px] sm:text-xs">
                                <thead>
                                  <tr className="bg-orange-50/50 border-b border-orange-200 text-orange-900">
                                    <th className="px-2 py-1">Farmer</th>
                                    <th className="px-2 py-1">Last Visit</th>
                                    <th className="px-2 py-1">Age</th>
                                    <th className="px-2 py-1 text-right">
                                      Feed (g/b)
                                    </th>
                                    <th className="px-2 py-1 text-right">
                                      Std Feed
                                    </th>
                                    <th className="px-2 py-1 text-right">
                                      Weight
                                    </th>
                                    <th className="px-2 py-1 text-right">
                                      Std Wt
                                    </th>
                                    <th className="px-2 py-1 text-right font-bold text-red-600">
                                      Less than Std
                                    </th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-orange-100">
                                  {list.map((v: any) => (
                                    <tr
                                      key={v.id}
                                      className="hover:bg-orange-50/40 leading-tight"
                                    >
                                      <td className="px-2 py-0.5 font-bold text-slate-800 whitespace-nowrap">
                                        {v.batch_label ||
                                          v.farmers?.farmer_name}
                                      </td>
                                      <td className="px-2 py-0.5 text-slate-600 whitespace-nowrap">
                                        {v.lastVisitDate || "-"}
                                      </td>
                                      <td className="px-2 py-0.5 text-slate-700 whitespace-nowrap">
                                        {v.ageAtVisit}d
                                      </td>
                                      <td className="px-2 py-0.5 text-right font-semibold text-slate-800 whitespace-nowrap">
                                        {v.actualFeed > 0
                                          ? `${v.actualFeed}g`
                                          : "-"}
                                      </td>
                                      <td className="px-2 py-0.5 text-right text-slate-600 whitespace-nowrap">
                                        {v.stdFeed > 0 ? `${v.stdFeed}g` : "-"}
                                      </td>
                                      <td className="px-2 py-0.5 text-right font-semibold whitespace-nowrap">
                                        {v.actual}g
                                      </td>
                                      <td className="px-2 py-0.5 text-right text-slate-600 whitespace-nowrap">
                                        {v.stdWt}g
                                      </td>
                                      <td className="px-2 py-0.5 text-right font-bold text-red-600 whitespace-nowrap">
                                        -{v.lessBy.toFixed(0)}g
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

                  {/* SECTION 4: LAYER BELOW STANDARD EGG PRODUCTION */}
                  {Object.keys(lowEggByAgent).length > 0 && (
                    <div className="space-y-1.5 pt-1">
                      <h4 className="font-bold text-yellow-700 text-xs sm:text-sm uppercase tracking-wider border-b pb-0.5">
                        Layer — Below Standard Egg Production
                      </h4>
                      {Object.entries(lowEggByAgent).map(
                        ([agentName, list]) => (
                          <div
                            key={agentName}
                            className="bg-yellow-50/30 border border-yellow-200 rounded-xl overflow-hidden shadow-xs"
                          >
                            <div className="bg-yellow-50 px-3 py-1 font-bold text-yellow-900 border-b border-yellow-200">
                              Agent:{" "}
                              <span className="text-emerald-700">
                                {agentName}
                              </span>
                            </div>
                            <div className="overflow-x-auto">
                              <table className="w-full text-left border-collapse text-[11px] sm:text-xs">
                                <thead>
                                  <tr className="bg-yellow-50/50 border-b border-yellow-200 text-yellow-900">
                                    <th className="px-2 py-1">Farmer</th>
                                    <th className="px-2 py-1">Last Visit</th>
                                    <th className="px-2 py-1">Age</th>
                                    <th className="px-2 py-1 text-right">
                                      Feed (g/b)
                                    </th>
                                    <th className="px-2 py-1 text-right">
                                      Std Feed
                                    </th>
                                    <th className="px-2 py-1 text-right">
                                      Egg %
                                    </th>
                                    <th className="px-2 py-1 text-right">
                                      Std %
                                    </th>
                                    <th className="px-2 py-1 text-right font-bold text-red-600">
                                      Less than Std
                                    </th>
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-yellow-100">
                                  {list.map((v: any) => (
                                    <tr
                                      key={v.id}
                                      className="hover:bg-yellow-50/40 leading-tight"
                                    >
                                      <td className="px-2 py-0.5 font-bold text-slate-800 whitespace-nowrap">
                                        {v.batch_label ||
                                          v.farmers?.farmer_name}
                                      </td>
                                      <td className="px-2 py-0.5 text-slate-600 whitespace-nowrap">
                                        {v.lastVisitDate || "-"}
                                      </td>
                                      <td className="px-2 py-0.5 text-slate-700 whitespace-nowrap">
                                        {v.ageWeeks}w
                                      </td>
                                      <td className="px-2 py-0.5 text-right font-semibold text-slate-800 whitespace-nowrap">
                                        {v.actualFeed > 0
                                          ? `${v.actualFeed}g`
                                          : "-"}
                                      </td>
                                      <td className="px-2 py-0.5 text-right text-slate-600 whitespace-nowrap">
                                        {v.stdFeed > 0 ? `${v.stdFeed}g` : "-"}
                                      </td>
                                      <td className="px-2 py-0.5 text-right font-semibold whitespace-nowrap">
                                        {v.actual}%
                                      </td>
                                      <td className="px-2 py-0.5 text-right text-slate-600 whitespace-nowrap">
                                        {v.stdEgg}%
                                      </td>
                                      <td className="px-2 py-0.5 text-right font-bold text-red-600 whitespace-nowrap">
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

            <div className="px-4 py-2 border-t border-slate-200 bg-slate-50 flex justify-end">
              <button
                onClick={() => setIsModalOpen(false)}
                className="px-4 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-900 text-white text-xs font-semibold transition"
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