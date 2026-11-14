// ================= AquaVisitsPage.tsx =================
// NEW page — Aqua (fish) visits. Modelled on VetVisitsPage (Poultry Visit):
// same header/filter/history layout, farmer + pond/shed select with the
// blank-single-shed rule, compact text-xs look.
// Table: aqua_visits (see aqua-cattle-migration.sql).
// Permissions (see src/utils/accessControl.ts, Worker 1):
// - View / history / export / Map: everyone.
// - Add/Edit/Delete: officer only on their own Aqua category page
//   (canAddReport(role, assignedCategories, "Aqua")) AND only their own
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

export const AQUA_SPECIES = [
  "Rui",
  "Katla",
  "Mrigel",
  "Silver Carp",
  "Grass Carp",
  "Common Carp",
  "Tilapia",
  "Pangas",
  "Koi",
  "Shing",
  "Magur",
  "Pabda",
  "Other",
];

const EMPTY_VISIT_FORM: any = {
  farmer_id: "",
  shed_id: "",
  visit_date: new Date().toISOString().split("T")[0],
  species: "Rui",
  species_other: "",
  stocked: "",
  mortality: "",
  feed_kg: "",
  ph: "",
  dissolved_oxygen: "",
  temperature: "",
  remarks: "",
  agent_id: "",
  company_id: "",
};

const EMPTY_FARMER_FORM: any = {
  farmer_name: "",
  mobile: "",
  district: "",
  upazila: "",
  farm_category: "Aqua",
};

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

// ---------- Report validation (shared messages, identical across pages) ----------
// Mortality: mortality can never exceed total fish (stocked) for the visit.
// Date: no report dated after today; no report dated before the stocking /
// hatch date, when one is known. Each helper returns the alert message or null.
function validateMortality(mortality: number, cap: number): string | null {
  if (mortality > cap) {
    return "Mortality total bird er cheye beshi hote parbe na.";
  }
  return null;
}

function validateReportDate(
  reportDate: string,
  hatchDate: string | null | undefined
): string | null {
  const today = new Date().toISOString().split("T")[0];
  if (reportDate > today) {
    return "Agami date e report dewa jabe na.";
  }
  if (hatchDate && reportDate < hatchDate) {
    return "Hatch date er age report dewa jabe na.";
  }
  return null;
}

export default function AquaVisitsPage() {
  const { role } = useAuth();
  const roleStr = String(role || "viewer");
  const isFullAccess = isManagerOrAdmin(roleStr);

  const [currentUserId, setCurrentUserId] = useState<string | null>(null);
  const [assignedCategories, setAssignedCategories] = useState<string[]>([]);
  const [_assignedUpazilas, setAssignedUpazilas] = useState<string[]>([]);
  const [_assignedDistricts, setAssignedDistricts] = useState<string[]>([]);

  const [farmers, setFarmers] = useState<any[]>([]);
  const [agents, setAgents] = useState<any[]>([]);
  const [companies, setCompanies] = useState<any[]>([]);
  const [visits, setVisits] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);

  const [searchQuery, setSearchQuery] = useState("");
  const [filterSpecies, setFilterSpecies] = useState("all");
  const [filterDate, setFilterDate] = useState("");

  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [editingRow, setEditingRow] = useState<any>(null);
  const [form, setForm] = useState<any>({ ...EMPTY_VISIT_FORM });

  const [showFarmerForm, setShowFarmerForm] = useState(false);
  const [farmerForm, setFarmerForm] = useState<any>({ ...EMPTY_FARMER_FORM });
  // Extra ponds/sheds for the new farmer (first pond = farmer form itself).
  // Adding a 2nd pond prompts to name the previous one (blank => "Shed 1").
  const [extraSheds, setExtraSheds] = useState<any[]>([]);
  const [savingFarmer, setSavingFarmer] = useState(false);

  // Add/Edit/Delete gate for Aqua reports.
  const canAddAqua = canAddReport(roleStr, assignedCategories, "Aqua");
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
    setAssignedUpazilas(uUpa);
    setAssignedDistricts(uDist);
    setAssignedCategories(uCats);

    const [{ data: farmersData }, { data: agentsData }, { data: compData }, { data: visitData }] =
      await Promise.all([
        supabase.from("farmers").select("*").order("farmer_name"),
        supabase.from("agents").select("*").order("agent_name"),
        supabase.from("companies").select("*").order("company_name"),
        supabase
          .from("aqua_visits")
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

  // Date floor for the report form: this page stores NO stocking / hatch
  // date (aqua_visits has only visit_date; the earliest visit row is a
  // report, not a stocking date, so it must not be used as a floor).
  // Only the today-rule applies, so the floor is null and the date input
  // gets max={today} with no min.
  const todayStr = new Date().toISOString().split("T")[0];
  const floorDate: string | null = null;
  // Mortality cap for the visit being entered/edited: the stocked qty on
  // this row (no running stock is maintained on this page).
  const mortalityCap =
    form.stocked === "" || Number.isNaN(Number(form.stocked))
      ? undefined
      : Number(form.stocked);

  // When the farmer changes, default to their (only) shed.
  function onFarmerChange(farmerId: string) {
    const f = farmerById[farmerId];
    const sheds = f ? getShedsForFarmer(f) : [];
    setForm((prev: any) => ({
      ...prev,
      farmer_id: farmerId,
      shed_id: sheds.length === 1 ? String(sheds[0].id) : "",
    }));
  }

  const filteredVisits = useMemo(() => {
    return visits.filter((v: any) => {
      const f = farmerById[String(v.farmer_id)];
      const name = getVisitShedDisplayName(v, f).toLowerCase();
      if (
        searchQuery &&
        !name.includes(searchQuery.toLowerCase()) &&
        !String(v.remarks || "").toLowerCase().includes(searchQuery.toLowerCase())
      )
        return false;
      if (filterSpecies !== "all" && v.species !== filterSpecies) return false;
      if (filterDate && String(v.visit_date || "").slice(0, 10) !== filterDate)
        return false;
      return true;
    });
  }, [visits, farmerById, searchQuery, filterSpecies, filterDate]);

  const summary = useMemo(() => {
    let stocked = 0;
    let mortality = 0;
    let feed = 0;
    filteredVisits.forEach((v: any) => {
      stocked += Number(v.stocked || 0);
      mortality += Number(v.mortality || 0);
      feed += Number(v.feed_kg || 0);
    });
    return { count: filteredVisits.length, stocked, mortality, feed };
  }, [filteredVisits]);

  function openAddForm() {
    setEditingId(null);
    setEditingRow(null);
    setForm({ ...EMPTY_VISIT_FORM });
    setShowForm(true);
  }

  function openEditForm(row: any) {
    // Guard lives in the handler too, but keep the button from opening
    // the form for a row this user may not edit.
    if (!canAddAqua || !canEditRecord(roleStr, row.created_by, currentUserId))
      return;
    setEditingId(row.id);
    setEditingRow(row);
    setForm({
      farmer_id: String(row.farmer_id || ""),
      shed_id: row.shed_id != null ? String(row.shed_id) : "",
      visit_date: String(row.visit_date || "").slice(0, 10),
      species: AQUA_SPECIES.includes(row.species) ? row.species : "Other",
      species_other: AQUA_SPECIES.includes(row.species) ? "" : row.species || "",
      stocked: row.stocked ?? "",
      mortality: row.mortality ?? "",
      feed_kg: row.feed_kg ?? "",
      ph: row.ph ?? "",
      dissolved_oxygen: row.dissolved_oxygen ?? "",
      temperature: row.temperature ?? "",
      remarks: row.remarks || "",
      agent_id: row.agent_id != null ? String(row.agent_id) : "",
      company_id: row.company_id != null ? String(row.company_id) : "",
    });
    setShowForm(true);
  }

  async function handleSaveVisit(e: React.FormEvent) {
    e.preventDefault();
    // Handler-level guard (not just the button): Add/Edit/Delete on the
    // Aqua page needs the Aqua category, and edits need ownership.
    if (!canAddAqua) {
      alert("Apnar Aqua category assign kora nei — Aqua visit add/edit korte parben na.");
      return;
    }
    if (editingRow && !canEditRecord(roleStr, editingRow.created_by, currentUserId)) {
      alert("Ei visit apni add koren ni — edit korte parben na.");
      return;
    }
    if (!form.farmer_id) {
      alert("Farmer select korun.");
      return;
    }
    if (!form.visit_date) {
      alert("Visit date din.");
      return;
    }
    const species =
      form.species === "Other"
        ? String(form.species_other || "").trim()
        : form.species;
    if (!species) {
      alert("Species din.");
      return;
    }
    // Report validation — before any DB write, on both add and edit
    // (edit is detected by editingId / editingRow above).
    const dateError = validateReportDate(form.visit_date, floorDate);
    if (dateError) {
      alert(dateError);
      return;
    }
    const stockedNum = form.stocked === "" ? 0 : Number(form.stocked);
    const mortalityNum = form.mortality === "" ? 0 : Number(form.mortality);
    const mortalityError = validateMortality(mortalityNum, stockedNum);
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
      species,
      stocked: form.stocked === "" ? 0 : Number(form.stocked),
      mortality: form.mortality === "" ? 0 : Number(form.mortality),
      feed_kg: form.feed_kg === "" ? null : Number(form.feed_kg),
      ph: form.ph === "" ? null : Number(form.ph),
      dissolved_oxygen:
        form.dissolved_oxygen === "" ? null : Number(form.dissolved_oxygen),
      temperature: form.temperature === "" ? null : Number(form.temperature),
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
          .from("aqua_visits")
          .update(payload)
          .eq("id", editingId);
        if (error) throw error;
      } else {
        payload.created_by = currentUserId;
        const { error } = await supabase.from("aqua_visits").insert(payload);
        if (error) throw error;
      }
      setShowForm(false);
      setEditingId(null);
      setEditingRow(null);
      await loadInitialData();
    } catch (err: any) {
      console.error("Aqua visit save error:", err);
      alert(err?.message || "Save failed.");
    } finally {
      setSaving(false);
    }
  }

  async function handleDeleteVisit(row: any) {
    // Handler-level guard: Aqua category + ownership.
    if (!canAddAqua) {
      alert("Apnar Aqua category assign kora nei — delete korte parben na.");
      return;
    }
    if (!canEditRecord(roleStr, row.created_by, currentUserId)) {
      alert("Ei visit apni add koren ni — delete korte parben na.");
      return;
    }
    if (!window.confirm("Ei Aqua visit ta delete korben?")) return;
    const { error } = await supabase.from("aqua_visits").delete().eq("id", row.id);
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
      alert("Ei farmer er location save kora nei.");
      return;
    }
    window.open(
      `https://www.google.com/maps?q=${encodeURIComponent(`${lat},${lng}`)}`,
      "_blank"
    );
  }

  function exportCsv() {
    const header = [
      "Date",
      "Farmer",
      "Pond/Shed",
      "Species",
      "Stocked",
      "Mortality",
      "Feed (kg)",
      "pH",
      "DO",
      "Temp (C)",
      "Remarks",
    ];
    const rows = filteredVisits.map((v: any) => {
      const f = farmerById[String(v.farmer_id)];
      const sheds = f ? getShedsForFarmer(f) : [];
      const shed = getShedForVisit(v, f);
      const short = shed
        ? getShedShortName(f, shed, sheds.indexOf(shed), sheds.length)
        : "";
      const esc = (x: any) => `"${String(x ?? "").replace(/"/g, '""')}"`;
      return [
        String(v.visit_date || "").slice(0, 10),
        f?.farmer_name || "",
        short,
        v.species || "",
        v.stocked ?? "",
        v.mortality ?? "",
        v.feed_kg ?? "",
        v.ph ?? "",
        v.dissolved_oxygen ?? "",
        v.temperature ?? "",
        v.remarks || "",
      ]
        .map(esc)
        .join(",");
    });
    const csv = [header.map((h) => `"${h}"`).join(","), ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `aqua-visits-${new Date().toISOString().split("T")[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  // ---- Add Farmer (any officer; NOT category-gated) ----
  function addExtraShed() {
    setExtraSheds((prev: any[]) => {
      // 2nd pond added: prompt to name the previous one (blank => "Shed 1").
      const next = [...prev];
      if (next.length === 0) {
        const name = window.prompt(
          "Ager pond/shed er nam din (khali rakhle 'Shed 1' hobe):",
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
    if (!canAddFarmerHere) {
      alert("Apni farmer add korte parben na.");
      return;
    }
    if (!farmerForm.farmer_name.trim()) {
      alert("Farmer name din.");
      return;
    }
    setSavingFarmer(true);
    try {
      const { data: inserted, error } = await supabase
        .from("farmers")
        .insert({
          farmer_name: farmerForm.farmer_name.trim(),
          mobile: farmerForm.mobile.trim() || null,
          district: farmerForm.district.trim() || null,
          upazila: farmerForm.upazila.trim() || null,
          farm_category: "Aqua",
        })
        .select()
        .single();
      if (error) throw error;
      const farmerId = inserted.id;
      // First pond: blank shed_name (single-pond rule). Extra ponds use
      // the names from the form (2nd+ pond naming handled in addExtraShed).
      const shedsToInsert = [
        { farmer_id: farmerId, shed_name: "", farm_category: "Aqua" },
        ...extraSheds.map((s: any) => ({
          farmer_id: farmerId,
          shed_name: String(s.shed_name || "").trim(),
          farm_category: "Aqua",
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
    "inline-flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold px-3 py-1.5 rounded-lg";
  const btnGhost =
    "inline-flex items-center gap-1 bg-white border border-slate-300 text-slate-700 text-xs font-semibold px-3 py-1.5 rounded-lg";

  if (loading) {
    return (
      <div className="p-4 text-xs text-slate-500">Loading Aqua Visits...</div>
    );
  }

  return (
    <div className="p-3 md:p-4 space-y-3 max-w-6xl mx-auto">
      {/* Header */}
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <h1 className="text-base md:text-lg font-bold text-slate-800">
            Aqua Visit
          </h1>
          <p className="text-[11px] text-slate-500">
            Fish pond visits — species, stock, mortality, feed & water params
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
          {canAddAqua && (
            <button onClick={openAddForm} className={btnPrimary}>
              <Plus size={13} /> Add Aqua Visit
            </button>
          )}
        </div>
      </div>

      {/* Summary */}
      <div className="grid grid-cols-4 gap-2">
        {[
          { label: "Visits", value: summary.count },
          { label: "Stocked", value: summary.stocked },
          { label: "Mortality", value: summary.mortality },
          { label: "Feed (kg)", value: summary.feed },
        ].map((s) => (
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
        <select
          value={filterSpecies}
          onChange={(e) => setFilterSpecies(e.target.value)}
          className={`${inputCls} w-auto`}
        >
          <option value="all">All Species</option>
          {AQUA_SPECIES.filter((s) => s !== "Other").map((s) => (
            <option key={s} value={s}>
              {s}
            </option>
          ))}
        </select>
        <input
          type="date"
          value={filterDate}
          onChange={(e) => setFilterDate(e.target.value)}
          className={`${inputCls} w-auto`}
        />
        {(searchQuery || filterSpecies !== "all" || filterDate) && (
          <button
            onClick={() => {
              setSearchQuery("");
              setFilterSpecies("all");
              setFilterDate("");
            }}
            className="text-[11px] text-emerald-700 font-semibold"
          >
            Clear
          </button>
        )}
      </div>

      {/* History */}
      <div className="bg-white border border-slate-200 rounded-xl overflow-hidden">
        <div className="px-3 py-2 border-b border-slate-100 text-xs font-bold text-slate-700">
          Visit History ({filteredVisits.length})
        </div>
        {filteredVisits.length === 0 ? (
          <div className="p-4 text-xs text-slate-400 text-center">
            No aqua visits yet.
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-slate-50 text-slate-500 text-left">
                  <th className="px-2 py-1.5 font-semibold">Date</th>
                  <th className="px-2 py-1.5 font-semibold">Farmer / Pond</th>
                  <th className="px-2 py-1.5 font-semibold">Species</th>
                  <th className="px-2 py-1.5 font-semibold text-right">Stocked</th>
                  <th className="px-2 py-1.5 font-semibold text-right">Mortality</th>
                  <th className="px-2 py-1.5 font-semibold text-right">Feed (kg)</th>
                  <th className="px-2 py-1.5 font-semibold">Water (pH / DO / Temp)</th>
                  <th className="px-2 py-1.5 font-semibold">Remarks</th>
                  <th className="px-2 py-1.5 font-semibold text-right">Action</th>
                </tr>
              </thead>
              <tbody>
                {filteredVisits.map((v: any) => {
                  const f = farmerById[String(v.farmer_id)];
                  const canEdit =
                    canAddAqua &&
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
                      <td className="px-2 py-1.5">{v.species}</td>
                      <td className="px-2 py-1.5 text-right">{v.stocked ?? 0}</td>
                      <td className="px-2 py-1.5 text-right">{v.mortality ?? 0}</td>
                      <td className="px-2 py-1.5 text-right">{v.feed_kg ?? "-"}</td>
                      <td className="px-2 py-1.5 whitespace-nowrap">
                        {v.ph ?? "-"} / {v.dissolved_oxygen ?? "-"} /{" "}
                        {v.temperature ?? "-"}
                      </td>
                      <td className="px-2 py-1.5 max-w-[180px] truncate" title={v.remarks || ""}>
                        {v.remarks || "-"}
                      </td>
                      <td className="px-2 py-1.5">
                        <div className="flex items-center justify-end gap-1">
                          <button
                            onClick={() => openMapForFarmer(f)}
                            className="p-1 rounded-md text-teal-600 hover:bg-teal-50"
                            title="Map (everyone)"
                          >
                            <MapPin size={14} />
                          </button>
                          {canEdit && (
                            <>
                              <button
                                onClick={() => openEditForm(v)}
                                className="p-1 rounded-md text-blue-600 hover:bg-blue-50"
                                title="Edit"
                              >
                                <Pencil size={14} />
                              </button>
                              <button
                                onClick={() => handleDeleteVisit(v)}
                                className="p-1 rounded-md text-red-600 hover:bg-red-50"
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
                {editingId ? "Edit Aqua Visit" : "Add Aqua Visit"}
              </h2>
              <button
                onClick={() => setShowForm(false)}
                className="p-1 rounded-md text-slate-500 hover:bg-slate-100"
              >
                <X size={16} />
              </button>
            </div>
            <form onSubmit={handleSaveVisit} className="p-4 space-y-3">
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
                  <label className={labelCls}>Pond / Shed</label>
                  <select
                    value={form.shed_id}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, shed_id: e.target.value }))
                    }
                    className={inputCls}
                  >
                    <option value="">Select pond/shed...</option>
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
                  Pond: {selectedFarmer.farmer_name} (single pond — no separate
                  name)
                </div>
              )}
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className={labelCls}>Visit Date *</label>
                  <input
                    type="date"
                    value={form.visit_date}
                    min={floorDate || undefined}
                    max={todayStr}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, visit_date: e.target.value }))
                    }
                    className={inputCls}
                    required
                  />
                </div>
                <div>
                  <label className={labelCls}>Species *</label>
                  <select
                    value={form.species}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, species: e.target.value }))
                    }
                    className={inputCls}
                  >
                    {AQUA_SPECIES.map((s) => (
                      <option key={s} value={s}>
                        {s}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
              {form.species === "Other" && (
                <div>
                  <label className={labelCls}>Species (other) *</label>
                  <input
                    value={form.species_other}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, species_other: e.target.value }))
                    }
                    placeholder="Species name likhun"
                    className={inputCls}
                  />
                </div>
              )}
              <div className="grid grid-cols-3 gap-2">
                <div>
                  <label className={labelCls}>Stocked Qty</label>
                  <input
                    type="number"
                    min={0}
                    value={form.stocked}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, stocked: e.target.value }))
                    }
                    className={inputCls}
                  />
                </div>
                <div>
                  <label className={labelCls}>Mortality</label>
                  <input
                    type="number"
                    min={0}
                    max={mortalityCap}
                    value={form.mortality}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, mortality: e.target.value }))
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
              <div>
                <div className={labelCls}>Water Params (optional)</div>
                <div className="grid grid-cols-3 gap-2">
                  <input
                    type="number"
                    step="0.01"
                    value={form.ph}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, ph: e.target.value }))
                    }
                    placeholder="pH"
                    className={inputCls}
                  />
                  <input
                    type="number"
                    step="0.01"
                    value={form.dissolved_oxygen}
                    onChange={(e) =>
                      setForm((p: any) => ({
                        ...p,
                        dissolved_oxygen: e.target.value,
                      }))
                    }
                    placeholder="DO (mg/L)"
                    className={inputCls}
                  />
                  <input
                    type="number"
                    step="0.1"
                    value={form.temperature}
                    onChange={(e) =>
                      setForm((p: any) => ({ ...p, temperature: e.target.value }))
                    }
                    placeholder="Temp (°C)"
                    className={inputCls}
                  />
                </div>
              </div>
              <div className="grid grid-cols-2 gap-2">
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
              <h2 className="text-sm font-bold text-slate-800">Add Farmer (Aqua)</h2>
              <button
                onClick={() => setShowFarmerForm(false)}
                className="p-1 rounded-md text-slate-500 hover:bg-slate-100"
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
                  <label className={labelCls}>Mobile</label>
                  <input
                    value={farmerForm.mobile}
                    onChange={(e) =>
                      setFarmerForm((p: any) => ({ ...p, mobile: e.target.value }))
                    }
                    className={inputCls}
                  />
                </div>
                <div>
                  <label className={labelCls}>Category</label>
                  <input value="Aqua" disabled className={`${inputCls} bg-slate-50`} />
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
                    Extra Pond {i + 2}
                  </div>
                  <div>
                    <label className={labelCls}>Pond / Shed Name</label>
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
                      placeholder="Shed 2 or Bridge par"
                      className={inputCls}
                    />
                  </div>
                </div>
              ))}
              <button
                type="button"
                onClick={addExtraShed}
                className="text-[11px] font-semibold text-emerald-700"
              >
                + Add Pond / Shed
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
