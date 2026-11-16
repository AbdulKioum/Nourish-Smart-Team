// ================= CattleVisitsPage.tsx =================
// NEW page — Cattle visits, two tabs: Dairy and Fattening. Modelled on
// VetVisitsPage (Poultry Visit): same header/filter/history layout,
// farmer + shed select with the blank-single-shed rule, compact text-xs look.
// Table: cattle_visits (see aqua-cattle-migration.sql), one row per visit
// with visit_type 'dairy' | 'fattening'.
// Permissions (see src/utils/accessControl.ts, Worker 1):
// - View / history / export / Map: everyone.
// - Add/Edit/Delete: officer only on their own Cattle category page
//   (canAddReport(role, assignedCategories, "Cattle")) AND only their own
//   rows (canEditRecord — old rows with created_by NULL are editable by
//   officers of the assigned area; the save stamps created_by).
// - Farmers can be added by any officer (not category-gated).
import React, { useEffect, useMemo, useState } from "react";
import { supabase } from "../../lib/supabase";
import { useAuth } from "../../context/AuthContext";
import {
  isManagerOrAdmin,
  canEditRecord,
  canAddReport,
  canAddFarmer,
} from "../../utils/accessControl";
import { MapPin, Plus, Trash2, Pencil, X, Search, Download } from "lucide-react";

type VisitType = "dairy" | "fattening";

const EMPTY_VISIT_FORM: any = {
  farmer_id: "",
  shed_id: "",
  visit_date: new Date().toISOString().split("T")[0],
  visit_type: "dairy" as VisitType,
  milk_litres: "",
  weight_kg: "",
  weight_gain_kg: "",
  feed_kg: "",
  mortality: "",
  health_note: "",
  remarks: "",
  agent_id: "",
  company_id: "",
};

const EMPTY_FARMER_FORM: any = {
  farmer_name: "",
  phone: "",
  district: "",
  upazila: "",
  farm_category: "Cattle",
};

// ---------- Report validation (shared messages, identical on every page) ----------
// Mortality cap = the selected shed's cattle on hand: dairy_cow + bull +
// calf_heifer from farmer_sheds (loaded with select("*") in loadInitialData;
// legacy single-shed fallback copies the same fields off the farmer row,
// mirroring FarmersPage / VisitsPage). This page never subtracts mortality
// from a running stock — the save only writes the cattle_visits row — so
// there is NO old-mortality add-back on edit: the cap is the shed stock
// as-is, on both add and edit. When no stock is recorded for the shed
// (sum is 0 / no shed selected) the cap is unknown and the mortality check
// is skipped rather than blocking a legitimate save.
// Date floor: this page stores NO stocking / hatch date (cattle_visits has
// no hatch_date, farmers / farmer_sheds have no stocking date), so the
// floor is null and only the today-rule applies.
const CATTLE_STOCK_FIELDS = ["dairy_cow", "bull", "calf_heifer"];

function getShedCattleStock(shed: any): number {
  return CATTLE_STOCK_FIELDS.reduce(
    (sum, k) => sum + Number(shed?.[k] || 0),
    0
  );
}

function validateReportDate(
  reportDate: string,
  hatchDate: string | null | undefined
): string | null {
  const d = String(reportDate || "").slice(0, 10);
  if (!d) return null; // empty date is caught by the required-field check
  const today = new Date().toISOString().split("T")[0];
  if (d > today) return "You cannot add a report for a future date.";
  const floor = hatchDate ? String(hatchDate).slice(0, 10) : "";
  if (floor && d < floor) return "Report date cannot be before the hatch date.";
  return null;
}

function validateMortality(
  mortality: number,
  cap: number | null | undefined
): string | null {
  if (cap == null) return null; // no stock on record — nothing to cap against
  if (Number(mortality || 0) > cap)
    return "Mortality cannot be more than total birds.";
  return null;
}

// ---------- Multi-shed helpers ----------
// Identical to FarmersPage / VisitsPage / DocTrackingPage. Stored shed_name
// is the SHORT label only ("Shed 1", "Bridge par"). A farmer with a single
// shed (no shed explicitly added) has NO shed name: the Shed column is
// blank and the full name is just the farmer name.
function getShedsForFarmer(farmer: any): any[] {
  const list = farmer?.farmer_sheds || farmer?.sheds || [];
  if (Array.isArray(list) && list.length > 0) return list;
  const shed: any = {
    id: `legacy-${farmer?.id}`,
    farmer_id: farmer?.id,
    shed_name: "",
    farm_category: farmer?.farm_category || "",
    latitude: farmer?.latitude ?? null,
    longitude: farmer?.longitude ?? null,
  };
  CATTLE_STOCK_FIELDS.forEach((k) => {
    shed[k] = Number(farmer?.[k] || 0);
  });
  return [shed];
}

function getShedShortName(
  farmer: any,
  shed: any,
  index: number = 0,
  totalSheds: number = 1
): string {
  const farmerName = String(farmer?.farmer_name || "").trim() || "Farmer";
  const raw = String(shed?.shed_name || "").trim();
  const lowerRaw = raw.toLowerCase();
  const lowerFarmer = farmerName.toLowerCase();
  const firstName = lowerFarmer.split(/\s+/)[0] || "";
  let rest = raw;
  const stripPrefix = (prefix: string): boolean => {
    if (prefix && lowerRaw.startsWith(prefix)) {
      const next = raw.charAt(prefix.length);
      if (next === " " || next === "-") {
        rest = raw.slice(prefix.length).replace(/^[\s-]+/, "").trim();
        return true;
      }
    }
    return false;
  };
  if (rest && rest.toLowerCase() !== lowerFarmer) {
    if (!stripPrefix(lowerFarmer) && firstName) stripPrefix(firstName);
  } else {
    rest = "";
  }
  const autoMatch = rest.match(/^shed\s*-?\s*(\d+)$/i);
  if ((totalSheds ?? 1) <= 1) {
    if (!rest || autoMatch) return "";
    return rest;
  }
  if (!rest) return `Shed ${index + 1}`;
  if (autoMatch) return `Shed ${autoMatch[1]}`;
  return rest;
}

function getShedDisplayName(
  farmer: any,
  shed: any,
  index: number = 0,
  totalSheds: number = 1
): string {
  const farmerName = String(farmer?.farmer_name || "").trim() || "Farmer";
  const short = getShedShortName(farmer, shed, index, totalSheds);
  return short ? `${farmerName} - ${short}` : farmerName;
}

function getShedForVisit(visit: any, farmer: any): any | null {
  if (!farmer) return null;
  const sheds = getShedsForFarmer(farmer);
  if (sheds.length === 0) return null;
  const sid = visit?.shed_id;
  if (sid !== null && sid !== undefined && String(sid) !== "") {
    return sheds.find((s: any) => String(s.id) === String(sid)) || null;
  }
  return sheds.length === 1 ? sheds[0] : null;
}

function getVisitShedDisplayName(visit: any, farmer: any): string {
  const farmerName =
    farmer?.farmer_name || visit?.farmers?.farmer_name || "Farmer";
  if (!farmer) return farmerName;
  const sheds = getShedsForFarmer(farmer);
  const shed = getShedForVisit(visit, farmer);
  if (!shed) return farmerName;
  return getShedDisplayName(farmer, shed, sheds.indexOf(shed), sheds.length);
}

function parseList(v: any): string[] {
  if (Array.isArray(v)) return v.map((x: any) => String(x));
  if (typeof v === "string" && v.trim()) {
    try {
      const p = JSON.parse(v);
      if (Array.isArray(p)) return p.map((x: any) => String(x));
    } catch {
      return v.split(",").map((s) => s.trim()).filter(Boolean);
    }
  }
  return [];
}

export default function CattleVisitsPage() {
  const { role } = useAuth();
  const roleStr = String(role || "viewer");
  const isFullAccess = isManagerOrAdmin(roleStr);

  const [currentUserId, setCurrentUserId] = useState<string | null>(null);
  const [assignedCategories, setAssignedCategories] = useState<string[]>([]);

  const [farmers, setFarmers] = useState<any[]>([]);
  const [agents, setAgents] = useState<any[]>([]);
  const [companies, setCompanies] = useState<any[]>([]);
  const [visits, setVisits] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);

  const [activeTab, setActiveTab] = useState<VisitType>("dairy");
  const [searchQuery, setSearchQuery] = useState("");
  const [filterDate, setFilterDate] = useState("");

  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [editingRow, setEditingRow] = useState<any>(null);
  const [form, setForm] = useState<any>({ ...EMPTY_VISIT_FORM });

  const [showFarmerForm, setShowFarmerForm] = useState(false);
  const [farmerForm, setFarmerForm] = useState<any>({ ...EMPTY_FARMER_FORM });
  // Extra sheds for the new farmer (first shed = farmer form itself).
  // Adding a 2nd shed prompts to name the previous one (blank => "Shed 1").
  const [extraSheds, setExtraSheds] = useState<any[]>([]);
  const [savingFarmer, setSavingFarmer] = useState(false);

  // Add/Edit/Delete gate for Cattle reports.
  const canAddCattle = canAddReport(roleStr, assignedCategories, "Cattle");
  // Farmer add is NOT category-gated: any officer may add a farmer.
  const canAddFarmerHere = canAddFarmer(roleStr);

  useEffect(() => {
    loadInitialData();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  async function loadInitialData() {
    setLoading(true);
    const {
      data: { user },
    } = await supabase.auth.getUser();
    setCurrentUserId(user?.id || null);

    let uUpa: string[] = [];
    let uDist: string[] = [];
    let uCats: string[] = [];
    if (user) {
      // assigned_categories is added by the permissions migration (Worker 1);
      // fall back gracefully if the column is not there yet.
      let uProf: any = null;
      const { data, error } = await supabase
        .from("users")
        .select("full_name, email, upazilas, districts, assigned_categories")
        .eq("id", user.id)
        .single();
      if (!error && data) {
        uProf = data;
      } else {
        const retry = await supabase
          .from("users")
          .select("full_name, email, upazilas, districts")
          .eq("id", user.id)
          .single();
        uProf = retry.data || null;
      }
      if (uProf) {
        uUpa = parseList(uProf.upazilas);
        uDist = parseList(uProf.districts);
        uCats = parseList(uProf.assigned_categories);
      }
    }
    setAssignedCategories(uCats);

    const [{ data: farmersData }, { data: agentsData }, { data: compData }, { data: visitData }] =
      await Promise.all([
        supabase.from("farmers").select("*").order("farmer_name"),
        supabase.from("agents").select("*").order("agent_name"),
        supabase.from("companies").select("*").order("company_name"),
        supabase
          .from("cattle_visits")
          .select("*")
          .order("visit_date", { ascending: false }),
      ]);

    let farmersList = farmersData || [];
    try {
      const { data: shedData, error: shedError } = await supabase
        .from("farmer_sheds")
        .select("*");
      if (!shedError && shedData) {
        const shedsByFarmer: Record<string, any[]> = {};
        shedData.forEach((s: any) => {
          const k = String(s.farmer_id);
          if (!shedsByFarmer[k]) shedsByFarmer[k] = [];
          shedsByFarmer[k].push(s);
        });
        farmersList.forEach((f: any) => {
          if (shedsByFarmer[String(f.id)])
            f.farmer_sheds = shedsByFarmer[String(f.id)];
        });
      }
    } catch {
      // farmer_sheds table not created yet — legacy fallback applies.
    }

    if (!isFullAccess) {
      farmersList = farmersList.filter((f: any) => {
        const matchUpa = uUpa.length === 0 || uUpa.includes(f.upazila);
        const matchDist = uDist.length === 0 || uDist.includes(f.district);
        return matchUpa && matchDist;
      });
    }
    setFarmers(farmersList);
    setAgents(agentsData || []);
    setCompanies(compData || []);
    setVisits(visitData || []);
    setLoading(false);
  }

  const farmerById = useMemo(() => {
    const m: Record<string, any> = {};
    farmers.forEach((f: any) => (m[String(f.id)] = f));
    return m;
  }, [farmers]);

  const selectedFarmer = form.farmer_id
    ? farmerById[String(form.farmer_id)]
    : null;
  const selectedFarmerSheds = selectedFarmer
    ? getShedsForFarmer(selectedFarmer)
    : [];

  // Validation inputs for the report form (see helpers above).
  const todayStr = new Date().toISOString().split("T")[0];
  // No stocking / hatch date is stored for cattle, so there is no floor date.
  const floorDate: string | null = null;
  const selectedShed =
    selectedFarmerSheds.find(
      (s: any) => String(s.id) === String(form.shed_id)
    ) || (selectedFarmerSheds.length === 1 ? selectedFarmerSheds[0] : null);
  const selectedShedStock = selectedShed
    ? getShedCattleStock(selectedShed)
    : 0;
  // Cap is the shed's cattle on hand; null when no stock is recorded
  // (unknown — the mortality check is then skipped, not blocking).
  const mortalityCap: number | null =
    selectedShed && selectedShedStock > 0 ? selectedShedStock : null;

  function onFarmerChange(farmerId: string) {
    const f = farmerById[farmerId];
    const sheds = f ? getShedsForFarmer(f) : [];
    setForm((prev: any) => ({
      ...prev,
      farmer_id: farmerId,
      shed_id: sheds.length === 1 ? String(sheds[0].id) : "",
    }));
  }

  const tabVisits = useMemo(
    () => visits.filter((v: any) => (v.visit_type || "dairy") === activeTab),
    [visits, activeTab]
  );

  const filteredVisits = useMemo(() => {
    return tabVisits.filter((v: any) => {
      const f = farmerById[String(v.farmer_id)];
      const name = getVisitShedDisplayName(v, f).toLowerCase();
      if (
        searchQuery &&
        !name.includes(searchQuery.toLowerCase()) &&
        !String(v.remarks || "").toLowerCase().includes(searchQuery.toLowerCase()) &&
        !String(v.health_note || "").toLowerCase().includes(searchQuery.toLowerCase())
      )
        return false;
      if (filterDate && String(v.visit_date || "").slice(0, 10) !== filterDate)
        return false;
      return true;
    });
  }, [tabVisits, farmerById, searchQuery, filterDate]);

  const tabCounts = useMemo(() => {
    const dairy = visits.filter((v: any) => (v.visit_type || "dairy") === "dairy").length;
    const fattening = visits.filter((v: any) => v.visit_type === "fattening").length;
    return { dairy, fattening };
  }, [visits]);

  const summary = useMemo(() => {
    let milk = 0;
    let weight = 0;
    let gain = 0;
    let feed = 0;
    let mortality = 0;
    filteredVisits.forEach((v: any) => {
      milk += Number(v.milk_litres || 0);
      weight += Number(v.weight_kg || 0);
      gain += Number(v.weight_gain_kg || 0);
      feed += Number(v.feed_kg || 0);
      mortality += Number(v.mortality || 0);
    });
    return { count: filteredVisits.length, milk, weight, gain, feed, mortality };
  }, [filteredVisits]);

  function openAddForm() {
    setEditingId(null);
    setEditingRow(null);
    setForm({ ...EMPTY_VISIT_FORM, visit_type: activeTab });
    setShowForm(true);
  }

  function openEditForm(row: any) {
    if (!canAddCattle || !canEditRecord(roleStr, row.created_by, currentUserId))
      return;
    setEditingId(row.id);
    setEditingRow(row);
    setForm({
      farmer_id: String(row.farmer_id || ""),
      shed_id: row.shed_id != null ? String(row.shed_id) : "",
      visit_date: String(row.visit_date || "").slice(0, 10),
      visit_type: row.visit_type === "fattening" ? "fattening" : "dairy",
      milk_litres: row.milk_litres ?? "",
      weight_kg: row.weight_kg ?? "",
      weight_gain_kg: row.weight_gain_kg ?? "",
      feed_kg: row.feed_kg ?? "",
      mortality: row.mortality ?? "",
      health_note: row.health_note || "",
      remarks: row.remarks || "",
      agent_id: row.agent_id != null ? String(row.agent_id) : "",
      company_id: row.company_id != null ? String(row.company_id) : "",
    });
    setShowForm(true);
  }

  async function handleSaveVisit(e: React.FormEvent) {
    e.preventDefault();
    if (saving) return;
    // Handler-level guard (not just the button): Add/Edit/Delete on the
    // Cattle page needs the Cattle category, and edits need ownership.
    if (!canAddCattle) {
      alert("Cattle is not assigned to you — you cannot add or edit Cattle visits.");
      return;
    }
    if (editingRow && !canEditRecord(roleStr, editingRow.created_by, currentUserId)) {
      alert("You did not add this visit — you cannot edit it.");
      return;
    }
    if (!form.farmer_id) {
      alert("Please select a farmer.");
      return;
    }
    if (!form.visit_date) {
      alert("Please enter the visit date.");
      return;
    }
    // Report validation — BEFORE any DB write (add and edit alike).
    const dateError = validateReportDate(form.visit_date, floorDate);
    if (dateError) {
      alert(dateError);
      return;
    }
    const mortalityError = validateMortality(
      form.mortality === "" ? 0 : Number(form.mortality),
      mortalityCap
    );
    if (mortalityError) {
      alert(mortalityError);
      return;
    }
    setSaving(true);
    const payload: any = {
      farmer_id: Number(form.farmer_id),
      shed_id:
        form.shed_id && !String(form.shed_id).startsWith("legacy-")
          ? Number(form.shed_id)
          : null,
      visit_date: form.visit_date,
      visit_type: form.visit_type,
      milk_litres:
        form.visit_type === "dairy" && form.milk_litres !== ""
          ? Number(form.milk_litres)
          : null,
      weight_kg:
        form.visit_type === "fattening" && form.weight_kg !== ""
          ? Number(form.weight_kg)
          : null,
      weight_gain_kg:
        form.visit_type === "fattening" && form.weight_gain_kg !== ""
          ? Number(form.weight_gain_kg)
          : null,
      feed_kg: form.feed_kg === "" ? null : Number(form.feed_kg),
      mortality: form.mortality === "" ? 0 : Number(form.mortality),
      health_note: String(form.health_note || "").trim(),
      remarks: String(form.remarks || "").trim(),
      agent_id: form.agent_id ? Number(form.agent_id) : null,
      company_id: form.company_id ? Number(form.company_id) : null,
    };
    try {
      if (editingId) {
        // Old rows with created_by NULL get stamped with the editor, so
        // after this edit only the creator (or manager/super admin) can
        // edit/delete it.
        if (!editingRow?.created_by && currentUserId) {
          payload.created_by = currentUserId;
        }
        const { error } = await supabase
          .from("cattle_visits")
          .update(payload)
          .eq("id", editingId);
        if (error) throw error;
      } else {
        payload.created_by = currentUserId;
        const { error } = await supabase.from("cattle_visits").insert(payload);
        if (error) throw error;
      }
      setShowForm(false);
      setEditingId(null);
      setEditingRow(null);
      await loadInitialData();
    } catch (err: any) {
      console.error("Cattle visit save error:", err);
      alert(err?.message || "Save failed.");
    } finally {
      setSaving(false);
    }
  }

  async function handleDeleteVisit(row: any) {
    // Handler-level guard: Cattle category + ownership.
    if (!canAddCattle) {
      alert("Cattle is not assigned to you — you cannot delete.");
      return;
    }
    if (!canEditRecord(roleStr, row.created_by, currentUserId)) {
      alert("You did not add this visit — you cannot delete it.");
      return;
    }
    if (!window.confirm("Delete this Cattle visit?")) return;
    const { error } = await supabase.from("cattle_visits").delete().eq("id", row.id);
    if (error) {
      alert(error.message || "Delete failed.");
      return;
    }
    await loadInitialData();
  }

  function openMapForFarmer(farmer: any) {
    // Map is for EVERYONE. Uses the farmer's (or shed's) stored location.
    const sheds = farmer ? getShedsForFarmer(farmer) : [];
    const s = sheds.find((x: any) => x.latitude && x.longitude) || sheds[0];
    const lat = s?.latitude ?? farmer?.latitude;
    const lng = s?.longitude ?? farmer?.longitude;
    if (lat == null || lng == null || lat === "" || lng === "") {
      alert("No location saved for this farmer.");
      return;
    }
    window.open(
      `https://www.google.com/maps?q=${encodeURIComponent(`${lat},${lng}`)}`,
      "_blank"
    );
  }

  function exportCsv() {
    const header =
      activeTab === "dairy"
        ? ["Date", "Farmer", "Shed", "Milk (L)", "Feed (kg)", "Mortality/Cull", "Health/Vaccine Note", "Remarks"]
        : ["Date", "Farmer", "Shed", "Weight (kg)", "Weight Gain (kg)", "Feed (kg)", "Mortality/Cull", "Health/Vaccine Note", "Remarks"];
    const rows = filteredVisits.map((v: any) => {
      const f = farmerById[String(v.farmer_id)];
      const sheds = f ? getShedsForFarmer(f) : [];
      const shed = getShedForVisit(v, f);
      const short = shed
        ? getShedShortName(f, shed, sheds.indexOf(shed), sheds.length)
        : "";
      const esc = (x: any) => `"${String(x ?? "").replace(/"/g, '""')}"`;
      const common = [
        String(v.visit_date || "").slice(0, 10),
        f?.farmer_name || "",
        short,
      ];
      const rest =
        activeTab === "dairy"
          ? [v.milk_litres ?? "", v.feed_kg ?? "", v.mortality ?? "", v.health_note || "", v.remarks || ""]
          : [v.weight_kg ?? "", v.weight_gain_kg ?? "", v.feed_kg ?? "", v.mortality ?? "", v.health_note || "", v.remarks || ""];
      return [...common, ...rest].map(esc).join(",");
    });
    const csv = [header.map((h) => `"${h}"`).join(","), ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `cattle-${activeTab}-visits-${new Date().toISOString().split("T")[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  // ---- Add Farmer (any officer; NOT category-gated) ----
  function addExtraShed() {
    setExtraSheds((prev: any[]) => {
      // 2nd shed added: prompt to name the previous one (blank => "Shed 1").
      const next = [...prev];
      if (next.length === 0) {
        const name = window.prompt(
          "Name the previous shed (leave empty for 'Shed 1'):",
          ""
        );
        next.push({
          shed_name: (name || "").trim() || "Shed 1",
          latitude: "",
          longitude: "",
        });
      } else {
        next.push({ shed_name: "", latitude: "", longitude: "" });
      }
      return next;
    });
  }

  async function handleSaveFarmer(e: React.FormEvent) {
    e.preventDefault();
    if (savingFarmer) return;
    if (!canAddFarmerHere) {
      alert("You do not have permission to add farmers.");
      return;
    }
    if (!farmerForm.farmer_name.trim()) {
      alert("Please enter the farmer name.");
      return;
    }
    setSavingFarmer(true);
    try {
      const { data: inserted, error } = await supabase
        .from("farmers")
        .insert({
          farmer_name: farmerForm.farmer_name.trim(),
          phone: farmerForm.phone.trim() || null,
          district: farmerForm.district.trim() || null,
          upazila: farmerForm.upazila.trim() || null,
          farm_category: "Cattle",
        })
        .select()
        .single();
      if (error) throw error;
      const farmerId = inserted.id;
      const shedsToInsert = [
        { farmer_id: farmerId, shed_name: "", farm_category: "Cattle" },
        ...extraSheds.map((s: any) => ({
          farmer_id: farmerId,
          shed_name: String(s.shed_name || "").trim(),
          farm_category: "Cattle",
          latitude: s.latitude === "" ? null : Number(s.latitude),
          longitude: s.longitude === "" ? null : Number(s.longitude),
        })),
      ];
      try {
        await supabase.from("farmer_sheds").insert(shedsToInsert);
      } catch {
        // farmer_sheds table not created yet — farmer row alone is enough.
      }
      setShowFarmerForm(false);
      setFarmerForm({ ...EMPTY_FARMER_FORM });
      setExtraSheds([]);
      await loadInitialData();
      setForm((prev: any) => ({ ...prev, farmer_id: String(farmerId) }));
    } catch (err: any) {
      console.error("Add farmer error:", err);
      alert(err?.message || "Farmer save failed.");
    } finally {
      setSavingFarmer(false);
    }
  }

  const inputCls =
    "w-full border border-slate-300 rounded-lg px-2 py-1.5 text-xs bg-white focus:outline-none focus:ring-1 focus:ring-emerald-500";
  const labelCls = "block text-[11px] font-semibold text-slate-600 mb-0.5";
  const btnPrimary =
    "inline-flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3 py-1.5 rounded-lg cursor-pointer disabled:opacity-60";
  const btnGhost =
    "inline-flex items-center gap-1 bg-white border border-slate-300 text-slate-700 text-xs font-semibold px-3 py-1.5 rounded-lg cursor-pointer";

  const tabBtn = (active: boolean) =>
    `px-3 py-1.5 rounded-full text-xs font-semibold cursor-pointer ${
      active ? "bg-emerald-600 text-white" : "bg-white text-slate-600 border border-slate-300"
    }`;

  if (loading) {
    return (
      <div className="p-4 text-xs text-slate-500">Loading Cattle Visits...</div>
    );
  }

  // TEMP: page parked for now — delete this block to restore the full page.
  return (
    <div className="min-h-[70vh] p-3 md:p-4 max-w-6xl mx-auto flex items-center justify-center overflow-hidden">
      <style>{`
        @keyframes ct-float { 0%, 100% { transform: translateY(0px); } 50% { transform: translateY(-12px); } }
        @keyframes ct-bubble { 0% { transform: translateY(0) scale(1); opacity: 0.6; } 100% { transform: translateY(-110px) scale(1.5); opacity: 0; } }
        @keyframes ct-shimmer { 0% { background-position: -200% center; } 100% { background-position: 200% center; } }
        @keyframes ct-ring { 0% { transform: scale(0.9); opacity: 0.7; } 100% { transform: scale(1.6); opacity: 0; } }
      `}</style>
      <div
        className="relative w-full max-w-lg text-center"
        style={{ perspective: "1000px" }}
      >
        {/* floating dots */}
        <span
          className="absolute left-[10%] bottom-6 w-3 h-3 rounded-full bg-lime-300/70"
          style={{ animation: "ct-bubble 4s ease-in infinite" }}
        />
        <span
          className="absolute left-[22%] bottom-2 w-2 h-2 rounded-full bg-green-300/70"
          style={{ animation: "ct-bubble 5s ease-in 1s infinite" }}
        />
        <span
          className="absolute right-[14%] bottom-8 w-4 h-4 rounded-full bg-amber-300/60"
          style={{ animation: "ct-bubble 4.5s ease-in 0.5s infinite" }}
        />
        <span
          className="absolute right-[26%] bottom-0 w-2 h-2 rounded-full bg-emerald-200/80"
          style={{ animation: "ct-bubble 3.5s ease-in 1.5s infinite" }}
        />
        {/* 3D floating card */}
        <div style={{ animation: "ct-float 5s ease-in-out infinite" }}>
          <div
            className="relative rounded-3xl bg-gradient-to-br from-lime-500 via-green-500 to-emerald-700 p-[2px] shadow-[0_30px_60px_-15px_rgba(22,163,74,0.55)]"
            style={{ transform: "rotateX(10deg)" }}
          >
            <div className="rounded-3xl bg-white/95 backdrop-blur px-6 py-10 sm:px-10">
              <div className="relative mx-auto w-20 h-20 mb-5">
                <span
                  className="absolute inset-0 rounded-full bg-green-400/40"
                  style={{ animation: "ct-ring 2s ease-out infinite" }}
                />
                <span className="relative flex items-center justify-center w-20 h-20 rounded-full bg-gradient-to-br from-lime-400 to-emerald-600 text-4xl shadow-lg shadow-green-500/40">
                  🐄
                </span>
              </div>
              <h1 className="text-xl sm:text-2xl font-extrabold tracking-tight text-slate-800">
                Cattle Visit
              </h1>
              <p
                className="mt-2 text-2xl sm:text-3xl font-black tracking-[0.2em] bg-gradient-to-r from-green-500 via-lime-500 to-green-500 bg-[length:200%_auto] bg-clip-text text-transparent"
                style={{ animation: "ct-shimmer 3s linear infinite" }}
              >
                COMING SOON
              </p>
              <p className="mt-3 text-sm text-slate-500">
                We&apos;re getting this page ready for you.
              </p>
              <div className="mt-6 flex items-center justify-center gap-1.5">
                <span className="w-2 h-2 rounded-full bg-green-500 animate-bounce" />
                <span
                  className="w-2 h-2 rounded-full bg-lime-500 animate-bounce"
                  style={{ animationDelay: "0.15s" }}
                />
                <span
                  className="w-2 h-2 rounded-full bg-emerald-500 animate-bounce"
                  style={{ animationDelay: "0.3s" }}
                />
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );

  return (
    <div className="p-3 md:p-4 space-y-3 max-w-6xl mx-auto">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <h1 className="text-base md:text-lg font-bold text-slate-800">
            Cattle Visit
          </h1>
          <p className="text-[11px] text-slate-500">
            Dairy & Fattening cattle visits — milk, weight, feed & health
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button onClick={exportCsv} className={btnGhost} title="Export CSV">
            <Download size={13} /> Export
          </button>
          {canAddFarmerHere && (
            <button
              onClick={() => setShowFarmerForm(true)}
              className={btnGhost}
              title="Add Farmer (any officer)"
            >
              <Plus size={13} /> Add Farmer
            </button>
          )}
          {canAddCattle && (
            <button onClick={openAddForm} className={btnPrimary}>
              <Plus size={13} /> Add Cattle Visit
            </button>
          )}
        </div>
      </div>

      {/* Dairy / Fattening tabs */}
      <div className="flex items-center gap-2">
        <button className={tabBtn(activeTab === "dairy")} onClick={() => setActiveTab("dairy")}>
          Dairy ({tabCounts.dairy})
        </button>
        <button
          className={tabBtn(activeTab === "fattening")}
          onClick={() => setActiveTab("fattening")}
        >
          Fattening ({tabCounts.fattening})
        </button>
      </div>

      {/* Summary */}
      <div className="grid grid-cols-4 gap-2">
        {(activeTab === "dairy"
          ? [
              { label: "Visits", value: summary.count },
              { label: "Milk (L)", value: summary.milk },
              { label: "Feed (kg)", value: summary.feed },
              { label: "Mortality/Cull", value: summary.mortality },
            ]
          : [
              { label: "Visits", value: summary.count },
              { label: "Weight (kg)", value: summary.weight },
              { label: "Weight Gain (kg)", value: summary.gain },
              { label: "Mortality/Cull", value: summary.mortality },
            ]
        ).map((s) => (
          <div
            key={s.label}
            className="bg-white border border-slate-200 rounded-xl px-3 py-2"
          >
            <div className="text-[11px] text-slate-500">{s.label}</div>
            <div className="text-base font-bold text-slate-800">{s.value}</div>
          </div>
        ))}
      </div>

      {/* Filters */}
      <div className="bg-white border border-slate-200 rounded-xl p-2.5 flex flex-wrap items-center gap-2">
        <div className="relative flex-1 min-w-[160px]">
          <Search
            size={13}
            className="absolute left-2 top-1/2 -translate-y-1/2 text-slate-400"
          />
          <input
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search farmer / remarks..."
            className={`${inputCls} pl-7`}
          />
        </div>
        <input
          type="date"
          value={filterDate}
          onChange={(e) => setFilterDate(e.target.value)}
          className={`${inputCls} w-auto`}
        />
        {(searchQuery || filterDate) && (
          <button
            onClick={() => {
              setSearchQuery("");
              setFilterDate("");
            }}
            className="text-[11px] text-emerald-700 font-semibold cursor-pointer"
          >
            Clear
          </button>
        )}
      </div>

      {/* History */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-3 py-2 border-b border-slate-100 text-xs font-bold text-slate-700">
          {activeTab === "dairy" ? "Dairy" : "Fattening"} History (
          {filteredVisits.length})
        </div>
        {filteredVisits.length === 0 ? (
          <div className="p-4 text-xs text-slate-400 text-center">
            No {activeTab} visits yet.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-slate-50 text-slate-500 text-left">
                  <th className="px-2 py-1.5 font-semibold">Date</th>
                  <th className="px-2 py-1.5 font-semibold">Farmer / Shed</th>
                  {activeTab === "dairy" ? (
                    <th className="px-2 py-1.5 font-semibold text-right">Milk (L)</th>
                  ) : (
                    <>
                      <th className="px-2 py-1.5 font-semibold text-right">Weight (kg)</th>
                      <th className="px-2 py-1.5 font-semibold text-right">Gain (kg)</th>
                    </>
                  )}
                  <th className="px-2 py-1.5 font-semibold text-right">Feed (kg)</th>
                  <th className="px-2 py-1.5 font-semibold text-right">Mortality/Cull</th>
                  <th className="px-2 py-1.5 font-semibold">Health / Vaccine Note</th>
                  <th className="px-2 py-1.5 font-semibold">Remarks</th>
                  <th className="px-2 py-1.5 font-semibold text-right">Action</th>
                </tr>
              </thead>
              <tbody>
                {filteredVisits.map((v: any) => {
                  const f = farmerById[String(v.farmer_id)];
                  const canEdit =
                    canAddCattle &&
                    canEditRecord(roleStr, v.created_by, currentUserId);
                  return (
                    <tr key={v.id} className="border-t border-slate-100">
                      <td className="px-2 py-1.5 whitespace-nowrap">
                        {String(v.visit_date || "").slice(0, 10)}
                      </td>
                      <td className="px-2 py-1.5">
                        <div className="font-semibold text-slate-700">
                          {getVisitShedDisplayName(v, f)}
                        </div>
                        <div className="text-[10px] text-slate-400">
                          {f?.upazila || ""}
                          {f?.district ? `, ${f.district}` : ""}
                        </div>
                      </td>
                      {activeTab === "dairy" ? (
                        <td className="px-2 py-1.5 text-right">{v.milk_litres ?? "-"}</td>
                      ) : (
                        <>
                          <td className="px-2 py-1.5 text-right">{v.weight_kg ?? "-"}</td>
                          <td className="px-2 py-1.5 text-right">{v.weight_gain_kg ?? "-"}</td>
                        </>
                      )}
                      <td className="px-2 py-1.5 text-right">{v.feed_kg ?? "-"}</td>
                      <td className="px-2 py-1.5 text-right">{v.mortality ?? 0}</td>
                      <td className="px-2 py-1.5 max-w-[160px] truncate" title={v.health_note || ""}>
                        {v.health_note || "-"}
                      </td>
                      <td className="px-2 py-1.5 max-w-[160px] truncate" title={v.remarks || ""}>
                        {v.remarks || "-"}
                      </td>
                      <td className="px-2 py-1.5">
                        <div className="flex items-center justify-end gap-1">
                          <button
                            onClick={() => openMapForFarmer(f)}
                            className="p-1 rounded-md text-teal-600 hover:bg-teal-50 cursor-pointer"
                            title="Map (everyone)"
                          >
                            <MapPin size={14} />
                          </button>
                          {canEdit && (
                            <>
                              <button
                                onClick={() => openEditForm(v)}
                                className="p-1 rounded-md text-blue-600 hover:bg-blue-50 cursor-pointer"
                                title="Edit"
                              >
                                <Pencil size={14} />
                              </button>
                              <button
                                onClick={() => handleDeleteVisit(v)}
                                className="p-1 rounded-md text-red-600 hover:bg-red-50 cursor-pointer"
                                title="Delete"
                              >
                                <Trash2 size={14} />
                              </button>
                            </>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Add / Edit Visit Modal */}
      {showForm && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-end md:items-center justify-center p-2">
          <div className="bg-white rounded-2xl w-full max-w-lg max-h-[92vh] overflow-y-auto">
            <div className="flex items-center justify-between px-4 py-2.5 border-b border-slate-100 sticky top-0 bg-white">
              <h2 className="text-sm font-bold text-slate-800">
                {editingId ? "Edit Cattle Visit" : "Add Cattle Visit"}
              </h2>
              <button
                onClick={() => setShowForm(false)}
                className="p-1 rounded-md text-slate-500 hover:bg-slate-100 cursor-pointer"
              >
                <X size={16} />
              </button>
            </div>
            <form onSubmit={handleSaveVisit} className="p-4 space-y-3">
              {/* Type switch inside the form too (editing a row keeps its type) */}
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  className={tabBtn(form.visit_type === "dairy")}
                  onClick={() =>
                    setForm((p: any) => ({ ...p, visit_type: "dairy" }))
                  }
                >
                  Dairy
                </button>
                <button
                  type="button"
                  className={tabBtn(form.visit_type === "fattening")}
                  onClick={() =>
                    setForm((p: any) => ({ ...p, visit_type: "fattening" }))
                  }
                >
                  Fattening
                </button>
              </div>
              <div>
                <label className={labelCls}>Farmer *</label>
                <select
                  value={form.farmer_id}
                  onChange={(e) => onFarmerChange(e.target.value)}
                  className={inputCls}
                  required
                >
                  <option value="">Select farmer...</option>
                  {farmers.map((f: any) => (
                    <option key={f.id} value={String(f.id)}>
                      {f.farmer_name}
                      {f.upazila ? ` (${f.upazila})` : ""}
                    </option>
                  ))}
                </select>
              </div>
              {selectedFarmer && selectedFarmerSheds.length > 1 && (
                <div>
                  <label className={labelCls}>Shed</label>
                  <select
                    value={form.shed_id}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, shed_id: e.target.value }))
                    }
                    className={inputCls}
                  >
                    <option value="">Select shed...</option>
                    {selectedFarmerSheds.map((s: any, i: number) => (
                      <option key={String(s.id)} value={String(s.id)}>
                        {getShedDisplayName(
                          selectedFarmer,
                          s,
                          i,
                          selectedFarmerSheds.length
                        )}
                      </option>
                    ))}
                  </select>
                </div>
              )}
              {selectedFarmer && selectedFarmerSheds.length === 1 && (
                <div className="text-[11px] text-slate-500">
                  Shed: {selectedFarmer.farmer_name} (single shed — no separate
                  name)
                </div>
              )}
              <div>
                <label className={labelCls}>Visit Date *</label>
                <input
                  type="date"
                  value={form.visit_date}
                  onChange={(e) =>
                    setForm((p: any) => ({ ...p, visit_date: e.target.value }))
                  }
                  className={inputCls}
                  max={todayStr}
                  required
                />
              </div>
              {form.visit_type === "dairy" ? (
                <div className="grid grid-cols-2 gap-2">
                  <div>
                    <label className={labelCls}>Milk Yield (litres)</label>
                    <input
                      type="number"
                      min={0}
                      step="0.1"
                      value={form.milk_litres}
                      onChange={(e) =>
                        setForm((p: any) => ({ ...p, milk_litres: e.target.value }))
                      }
                      className={inputCls}
                    />
                  </div>
                  <div>
                    <label className={labelCls}>Feed (kg)</label>
                    <input
                      type="number"
                      min={0}
                      step="0.01"
                      value={form.feed_kg}
                      onChange={(e) =>
                        setForm((p: any) => ({ ...p, feed_kg: e.target.value }))
                      }
                      className={inputCls}
                    />
                  </div>
                </div>
              ) : (
                <div className="grid grid-cols-3 gap-2">
                  <div>
                    <label className={labelCls}>Weight (kg)</label>
                    <input
                      type="number"
                      min={0}
                      step="0.1"
                      value={form.weight_kg}
                      onChange={(e) =>
                        setForm((p: any) => ({ ...p, weight_kg: e.target.value }))
                      }
                      className={inputCls}
                    />
                  </div>
                  <div>
                    <label className={labelCls}>Weight Gain (kg)</label>
                    <input
                      type="number"
                      step="0.1"
                      value={form.weight_gain_kg}
                      onChange={(e) =>
                        setForm((p: any) => ({
                          ...p,
                          weight_gain_kg: e.target.value,
                        }))
                      }
                      className={inputCls}
                    />
                  </div>
                  <div>
                    <label className={labelCls}>Feed (kg)</label>
                    <input
                      type="number"
                      min={0}
                      step="0.01"
                      value={form.feed_kg}
                      onChange={(e) =>
                        setForm((p: any) => ({ ...p, feed_kg: e.target.value }))
                      }
                      className={inputCls}
                    />
                  </div>
                </div>
              )}
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className={labelCls}>Mortality / Cull</label>
                  <input
                    type="number"
                    min={0}
                    max={mortalityCap ?? undefined}
                    value={form.mortality}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, mortality: e.target.value }))
                    }
                    className={inputCls}
                  />
                </div>
                <div>
                  <label className={labelCls}>Agent</label>
                  <select
                    value={form.agent_id}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, agent_id: e.target.value }))
                    }
                    className={inputCls}
                  >
                    <option value="">Select agent...</option>
                    {agents.map((a: any) => (
                      <option key={a.id} value={String(a.id)}>
                        {a.agent_name}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className={labelCls}>Company</label>
                  <select
                    value={form.company_id}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, company_id: e.target.value }))
                    }
                    className={inputCls}
                  >
                    <option value="">Select company...</option>
                    {companies.map((c: any) => (
                      <option key={c.id} value={String(c.id)}>
                        {c.company_name}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
              <div>
                <label className={labelCls}>Health / Vaccine Note</label>
                <input
                  value={form.health_note}
                  onChange={(e) =>
                    setForm((p: any) => ({ ...p, health_note: e.target.value }))
                  }
                  placeholder="Vaccine / treatment note"
                  className={inputCls}
                />
              </div>
              <div>
                <label className={labelCls}>Remarks</label>
                <textarea
                  value={form.remarks}
                  onChange={(e) =>
                    setForm((p: any) => ({ ...p, remarks: e.target.value }))
                  }
                  rows={2}
                  className={inputCls}
                />
              </div>
              <div className="flex justify-end gap-2 pt-1">
                <button
                  type="button"
                  onClick={() => setShowForm(false)}
                  className={btnGhost}
                >
                  Cancel
                </button>
                <button type="submit" disabled={saving} className={btnPrimary}>
                  {saving ? "Saving..." : editingId ? "Update" : "Save"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Add Farmer Modal */}
      {showFarmerForm && (
        <div className="fixed inset-0 z-50 bg-black/40 flex items-end md:items-center justify-center p-2">
          <div className="bg-white rounded-2xl w-full max-w-md max-h-[92vh] overflow-y-auto">
            <div className="flex items-center justify-between px-4 py-2.5 border-b border-slate-100 sticky top-0 bg-white">
              <h2 className="text-sm font-bold text-slate-800">Add Farmer (Cattle)</h2>
              <button
                onClick={() => setShowFarmerForm(false)}
                className="p-1 rounded-md text-slate-500 hover:bg-slate-100 cursor-pointer"
              >
                <X size={16} />
              </button>
            </div>
            <form onSubmit={handleSaveFarmer} className="p-4 space-y-3">
              <div>
                <label className={labelCls}>Farmer Name *</label>
                <input
                  value={farmerForm.farmer_name}
                  onChange={(e) =>
                    setFarmerForm((p: any) => ({ ...p, farmer_name: e.target.value }))
                  }
                  className={inputCls}
                  required
                />
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className={labelCls}>Phone</label>
                  <input
                    value={farmerForm.phone}
                    onChange={(e) =>
                      setFarmerForm((p: any) => ({ ...p, phone: e.target.value }))
                    }
                    className={inputCls}
                  />
                </div>
                <div>
                  <label className={labelCls}>Category</label>
                  <input value="Cattle" disabled className={`${inputCls} bg-slate-50`} />
                </div>
                <div>
                  <label className={labelCls}>District</label>
                  <input
                    value={farmerForm.district}
                    onChange={(e) =>
                      setFarmerForm((p: any) => ({ ...p, district: e.target.value }))
                    }
                    className={inputCls}
                  />
                </div>
                <div>
                  <label className={labelCls}>Upazila</label>
                  <input
                    value={farmerForm.upazila}
                    onChange={(e) =>
                      setFarmerForm((p: any) => ({ ...p, upazila: e.target.value }))
                    }
                    className={inputCls}
                  />
                </div>
              </div>
              {extraSheds.map((s: any, i: number) => (
                <div key={i} className="border border-slate-200 rounded-lg p-2 space-y-2">
                  <div className="text-[11px] font-bold text-slate-600">
                    Extra Shed {i + 2}
                  </div>
                  <div>
                    <label className={labelCls}>Shed Name</label>
                    <input
                      value={s.shed_name}
                      onChange={(e) => {
                        const v = e.target.value;
                        setExtraSheds((prev: any[]) =>
                          prev.map((x, xi) =>
                            xi === i ? { ...x, shed_name: v } : x
                          )
                        );
                      }}
                      placeholder="Shed 2 or Bridge side"
                      className={inputCls}
                    />
                  </div>
                </div>
              ))}
              <button
                type="button"
                onClick={addExtraShed}
                className="text-[11px] font-semibold text-emerald-700 cursor-pointer"
              >
                + Add Shed
              </button>
              <div className="flex justify-end gap-2 pt-1">
                <button
                  type="button"
                  onClick={() => setShowFarmerForm(false)}
                  className={btnGhost}
                >
                  Cancel
                </button>
                <button type="submit" disabled={savingFarmer} className={btnPrimary}>
                  {savingFarmer ? "Saving..." : "Save Farmer"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
