import React, { useEffect, useState, useRef } from "react";
import LocationPicker from "../../components/maps/LocationPicker";
import SearchableSelect from "../../components/common/SearchableSelect";
import { supabase } from "../../lib/supabase";
import {
  getAgents,
  createAgent,
  updateAgent,
  deleteAgent,
} from "../../services/agentService";
import { useAuth } from "../../context/AuthContext";
import {
  canEditRecord,
  canAddFarmer,
} from "../../utils/accessControl";
import {
  MapContainer,
  Marker,
  Popup,
  Tooltip,
  useMap,
  Polyline,
} from "react-leaflet";
import L from "leaflet";
import "leaflet/dist/leaflet.css";
import { saveAs } from "file-saver";
import { getBazars } from "../../services/bazarService";
import MapLayers from "../../components/maps/MapLayers";
import { createAgentMarker } from "../../utils/mapMarkers";
import ExcelJS from "exceljs";
import {
  MapPin,
  Plus,
  Search,
  RotateCcw,
  FileSpreadsheet,
  BarChart3,
  Map as MapIcon,
  Send,
  X,
  CheckCircle2,
  Navigation,
} from "lucide-react";
import { Filesystem, Directory } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";
import { Capacitor } from "@capacitor/core";
import { Geolocation } from "@capacitor/geolocation";
import { UserAvatar, renderNoteWithMentions } from "../../utils/avatarHelper";

delete (L.Icon.Default.prototype as any)._getIconUrl;

L.Icon.Default.mergeOptions({
  iconRetinaUrl:
    "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon-2x.png",
  iconUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png",
  shadowUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png",
});

const EMPTY_FORM = {
  company_id: "",
  agent_code: "",
  agent_name: "",
  proprietor: "",
  mobile: "",
  bazar: "",
  union_name: "",
  upazila: "",
  district: "",
  broiler_sale: "0",
  layer_sale: "0",
  sonali_sale: "0",
  cattle_sale: "0",
  fish_sale: "0",
  monthly_sale: "0",
  focus_item: "",
  latitude: "",
  longitude: "",
  status: "Active",
};

function calcDistanceKm(
  lat1: number,
  lon1: number,
  lat2: number,
  lon2: number
): number {
  const R = 6371;
  const dLat = ((lat2 - lat1) * Math.PI) / 180;
  const dLon = ((lon2 - lon1) * Math.PI) / 180;
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos((lat1 * Math.PI) / 180) *
      Math.cos((lat2 * Math.PI) / 180) *
      Math.sin(dLon / 2) *
      Math.sin(dLon / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function formatDistance(km: number): string {
  if (km < 1) return `${Math.round(km * 1000)} m`;
  return `${km.toFixed(2)} km`;
}

function parseList(val: any): string[] {
  if (!val) return [];
  if (Array.isArray(val)) return val.map(String).filter(Boolean);
  if (typeof val === "string") {
    try {
      const parsed = JSON.parse(val);
      if (Array.isArray(parsed)) return parsed.map(String).filter(Boolean);
    } catch {
      return val
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
    }
  }
  return [];
}

function getSaleBreakdown(agent: any) {
  const items = [
    { label: "Broiler", val: Number(agent?.broiler_sale || 0) },
    { label: "Layer", val: Number(agent?.layer_sale || 0) },
    { label: "Sonali", val: Number(agent?.sonali_sale || 0) },
    { label: "Cattle", val: Number(agent?.cattle_sale || 0) },
    { label: "Fish", val: Number(agent?.fish_sale || 0) },
  ];
  return items.filter((i) => i.val > 0);
}


const SALE_KEYS = ["broiler_sale", "sonali_sale", "layer_sale", "cattle_sale", "fish_sale"] as const;
const SALE_LABELS: Record<string, string> = {
  broiler_sale: "Broiler",
  sonali_sale: "Sonali",
  layer_sale: "Layer",
  cattle_sale: "Cattle",
  fish_sale: "Fish",
};
function linkSaleTotal(link: any): number {
  return SALE_KEYS.reduce((sum, k) => sum + Number(link?.[k] || 0), 0);
}

function emptyCompanyLink(companyId: any = ""): any {
  return {
    id: undefined,
    company_id: companyId ? String(companyId) : "",
    company_name: "",
    logo_url: "",
    broiler_sale: "0",
    sonali_sale: "0",
    layer_sale: "0",
    cattle_sale: "0",
    fish_sale: "0",
    monthly_sale: "0",
  };
}

// Linked companies for an agent. Exclusivity is DERIVED: 2+ links =
// non-exclusive. Falls back to the agent's single company_id + sale
// fields when agent_companies does not exist yet.
function getAgentCompanies(agent: any): any[] {
  const list = agent?.agent_companies || agent?.companies_list || [];
  if (Array.isArray(list) && list.length > 0) {
    return list.map((c: any) => ({
      ...c,
      company_id: String(c.company_id),
      monthly_sale: Number((c.monthly_sale ?? linkSaleTotal(c)) || 0),
    }));
  }
  if (!agent?.company_id) return [];
  return [
    {
      id: `legacy-${agent.id}`,
      agent_id: agent.id,
      company_id: String(agent.company_id),
      company_name: agent.companies?.company_name || "",
      logo_url: agent.companies?.logo_url || "",
      broiler_sale: agent.broiler_sale || 0,
      sonali_sale: agent.sonali_sale || 0,
      layer_sale: agent.layer_sale || 0,
      cattle_sale: agent.cattle_sale || 0,
      fish_sale: agent.fish_sale || 0,
      monthly_sale: agent.monthly_sale || 0,
    },
  ];
}

function isNonExclusive(agent: any): boolean {
  return getAgentCompanies(agent).length >= 2;
}

function getAgentTotalSale(agent: any): number {
  const links = getAgentCompanies(agent);
  if (links.length === 0) return Number(agent?.monthly_sale || 0);
  return links.reduce(
    (sum, l) => sum + Number((l.monthly_sale ?? linkSaleTotal(l)) || 0),
    0
  );
}

// Agent marker.
// EXCLUSIVE agents: exactly the original marker — createAgentMarker from
// ../../utils/mapMarkers with the agent's single (primary) company logo.
// (The MapMarker source file was not available in this workspace, so the
// original sizing comes from that import, not from code here.)
// NON-EXCLUSIVE agents only: multiple logos together, same name chip and
// same overall height as the original marker — no pin/emoji overlay, and
// logos are NOT shrunk (22px, the size a single logo uses).
function getAgentMarkerIcon(agent: any) {
  if (!isNonExclusive(agent)) {
    const links = getAgentCompanies(agent);
    const logoUrl = links[0]?.logo_url || agent.companies?.logo_url || "";
    return createAgentMarker(logoUrl, agent.agent_name || "");
  }
  return buildNonExclusiveAgentMarker(agent);
}

function buildNonExclusiveAgentMarker(agent: any) {
  const links = getAgentCompanies(agent);
  const logos = links.map((l: any) => l.logo_url || "").filter(Boolean);
  const shown = logos.slice(0, 3);
  const extra = logos.length - shown.length;
  const safeName = (agent.agent_name || "Agent")
    .replace(/</g, "&lt;")
    .slice(0, 18);
  const logosHtml =
    shown.length > 0
      ? `<div style="display:flex;align-items:center;justify-content:center;gap:2px;background:rgba(255,255,255,0.92);padding:2px 3px;border-radius:3px;box-shadow:0 1px 2px rgba(0,0,0,0.2);">${shown
          .map(
            (u) =>
              `<img src="${u}" style="width:22px;height:22px;object-fit:contain;" />`
          )
          .join("")}${
          extra > 0
            ? `<span style="font-size:10px;font-weight:700;color:#334155;">+${extra}</span>`
            : ""
        }</div>`
      : "";
  const html = `
    <div style="display:flex;flex-direction:column;align-items:center;transform:translateY(-14px);">
      ${logosHtml}
      <div style="margin-top:2px;background:rgba(252,244,236,0.95);border:1.5px solid #475569;color:#314158;font-size:10px;font-weight:700;padding:0px 5px;border-radius:4px;white-space:nowrap;box-shadow:0 1px 2px rgba(0,0,0,0.15);">
        ${safeName}
      </div>
    </div>
  `;
  const width = shown.length >= 3 ? 84 : 64;
  return L.divIcon({
    className: "custom-agent-div-icon",
    html,
    iconSize: [width, 56],
    iconAnchor: [width / 2, 46],
    popupAnchor: [0, -40],
  });
}

function isNourishName(name?: string) {
  return Boolean(name && name.toLowerCase().includes("nourish"));
}

function FitBounds({
  locations,
  skip,
}: {
  locations: { lat: number; lng: number }[];
  skip?: boolean;
}) {
  const map = useMap();

  useEffect(() => {
    if (skip) return;
    if (!locations || locations.length === 0) return;

    const bounds = L.latLngBounds(locations.map((loc) => [loc.lat, loc.lng]));
    map.fitBounds(bounds, { padding: [50, 50], maxZoom: 15 });
  }, [locations, map, skip]);

  return null;
}

function MapFocusController({
  agentPos,
  myPos,
  myLocationMode,
}: {
  agentPos: [number, number] | null;
  myPos: [number, number] | null;
  myLocationMode: boolean;
}) {
  const map = useMap();

  useEffect(() => {
    if (myPos && agentPos) {
      const bounds = L.latLngBounds([myPos, agentPos]);
      map.fitBounds(bounds, { padding: [65, 65], maxZoom: 16, animate: true });
    } else if (agentPos) {
      map.flyTo(agentPos, 17, { animate: true, duration: 1.0 });
    } else if (myLocationMode && myPos) {
      map.flyTo(myPos, 16, { animate: true, duration: 1.0 });
    }
  }, [agentPos, myPos, myLocationMode, map]);

  return null;
}

export default function AgentsPage() {
  const { role } = useAuth();
  const mapSectionRef = useRef<HTMLDivElement | null>(null);
  const agentMarkerRefs = useRef<Record<number, L.Marker | null>>({});
  const myLocationMarkerRef = useRef<L.Marker | null>(null);

  if (role === "viewer") {
    return <div className="p-6 text-center text-slate-500">Access Denied</div>;
  }

  const isFullAccess = role === "super_admin" || role === "manager";
  // Permissions (src/utils/accessControl.ts): manager / super admin can do
  // everything; officers edit/delete only their own rows (old rows with
  // created_by null are editable by area officers — the save stamps the
  // creator). Agent add is NOT category-gated (rule 5, like farmers).
  // View / Map stay open to everyone.
  const canAddAgentHere = canAddFarmer(role);
  const canEditAgent = (agent: any) =>
    canEditRecord(role, agent?.created_by, currentUserId);

  // User Profile & Assigned Areas
  const [currentUserId, setCurrentUserId] = useState("");
  const [currentUserName, setCurrentUserName] = useState("User");
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>([]);
  const [, setAssignedCategories] = useState<string[]>([]);

  // Master Data
  const [agents, setAgents] = useState<any[]>([]);
  const [companies, setCompanies] = useState<any[]>([]);
  const [districts, setDistricts] = useState<any[]>([]);
  const [upazilas, setUpazilas] = useState<any[]>([]);
  const [unions, setUnions] = useState<any[]>([]);
  const [bazars, setBazars] = useState<any[]>([]);

  // Form & Modals
  const [showForm, setShowForm] = useState(false);
  const [showMap, setShowMap] = useState(false);
  const [showAgentMap, setShowAgentMap] = useState(false);
  const [showSummary, setShowSummary] = useState(false);
  const [selectedAgent, setSelectedAgent] = useState<any>(null);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [form, setForm] = useState({ ...EMPTY_FORM });
  const [formCompanies, setFormCompanies] = useState<any[]>([]);
  const [formType, setFormType] = useState<"exclusive" | "nonexclusive">("exclusive");
  const [companyPickerId, setCompanyPickerId] = useState("");
  const [nonExPopupAgentId, setNonExPopupAgentId] = useState<number | null>(null);
  const [nonExPopupPos, setNonExPopupPos] = useState<{ top: number; left: number } | null>(null);
  const [showFilters, setShowFilters] = useState(true);
  const [selectedMapAgentId, setSelectedMapAgentId] = useState<number | null>(null);

  // Close the non-ex. popover with Escape.
  useEffect(() => {
    if (nonExPopupAgentId === null) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        setNonExPopupAgentId(null);
        setNonExPopupPos(null);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [nonExPopupAgentId]);

  // Location Picker Live Location State
  const [pickerCenterOverride, setPickerCenterOverride] = useState<
    [number, number] | null
  >(null);
  const [locatingInPicker, setLocatingInPicker] = useState(false);

  // Agent Tracking Notes State
  const [activeTrackingAgentId, setActiveTrackingAgentId] = useState<
    number | null
  >(null);
  const [agentNotes, setAgentNotes] = useState<any[]>([]);
  const [newNoteText, setNewNoteText] = useState("");
  const [loadingNotes, setLoadingNotes] = useState(false);
  const [savingNote, setSavingNote] = useState(false);
  const [savingAgent, setSavingAgent] = useState(false);

  // Filters
  const [search, setSearch] = useState("");
  const [companyFilter, setCompanyFilter] = useState("");
  const [districtFilter, setDistrictFilter] = useState("");
  const [upazilaFilter, setUpazilaFilter] = useState("");
  const [unionFilter, setUnionFilter] = useState("");
  const [bazarFilter, setBazarFilter] = useState("");
  const [exclusivityFilter, setExclusivityFilter] = useState("");

  const [filterUpazilas, setFilterUpazilas] = useState<any[]>([]);
  const [filterUnions, setFilterUnions] = useState<any[]>([]);
  const [filterBazars, setFilterBazars] = useState<any[]>([]);

  // Main Map States
  const [userLocation, setUserLocation] = useState<[number, number] | null>(
    null
  );
  const [nearestAgents, setNearestAgents] = useState<any[]>([]);
  const [selectedLocation, setSelectedLocation] = useState<
    [number, number] | null
  >(null);
  const [mapCenter] = useState<[number, number]>([24.75, 90.4]);
  const [myLocationMode, setMyLocationMode] = useState(false);

  // Check if selected company in form is Nourish
  const selectedFormCompanyObj = companies.find(
    (c) => String(c.id) === String(form.company_id)
  );
  const isFormCompanyNourish = isNourishName(
    selectedFormCompanyObj?.company_name
  );

  const [allUsers, setAllUsers] = useState<any[]>([]);

  useEffect(() => {
    supabase
      .from("users")
      .select("id, full_name, avatar_url")
      .then(({ data }) => {
        if (data) setAllUsers(data);
      });
  }, []);

  function getUserAvatar(userId?: string, userName?: string): string | null {
    const found = allUsers.find(
      (u) =>
        (userId && String(u.id) === String(userId)) ||
        (userName &&
          String(u.full_name || "").trim().toLowerCase() ===
            String(userName).trim().toLowerCase())
    );
    return found?.avatar_url || null;
  }

  // ==================================================
  // INITIALIZE USER & DATA
  // ==================================================
  async function loadData() {
    const {
      data: { user },
    } = await supabase.auth.getUser();

    let uDistricts: string[] = [];
    let uUpazilas: string[] = [];
    let uName = user?.email?.split("@")[0] || "Officer";

    if (user) {
      setCurrentUserId(user.id);
      // assigned_categories comes from the permissions migration;
      // fall back gracefully if the column is not there yet.
      let profile: any = null;
      {
        const { data, error } = await supabase
          .from("users")
          .select("full_name, email, role, districts, upazilas, assigned_categories")
          .eq("id", user.id)
          .single();
        if (!error && data) {
          profile = data;
        } else {
          const retry = await supabase
            .from("users")
            .select("full_name, email, role, districts, upazilas")
            .eq("id", user.id)
            .single();
          profile = retry.data || null;
        }
      }

      if (profile) {
        uName = profile.full_name || profile.email || uName;
        uDistricts = parseList(profile.districts);
        uUpazilas = parseList(profile.upazilas);
        setAssignedCategories(parseList(profile.assigned_categories));
      }
    }

    setCurrentUserName(uName);
    setAssignedUpazilas(uUpazilas);

    const [
      agentData,
      bazarData,
      { data: companyData },
      { data: districtData },
    ] = await Promise.all([
      getAgents(),
      getBazars(),
      supabase.from("companies").select("*").order("company_name"),
      supabase.from("districts").select("*").order("name"),
    ]);

    const allAgents = agentData || [];
    // Ownership: attach created_by when the column exists (the agent
    // service may not select it). A missing column must not break the
    // list — rows without a creator stay editable (rule 2, old data).
    try {
      const { data: creatorData, error: creatorErr } = await supabase
        .from("agents")
        .select("id, created_by");
      if (!creatorErr && creatorData) {
        const creatorById: Record<string, any> = {};
        (creatorData as any[]).forEach((r: any) => {
          creatorById[String(r.id)] = r.created_by;
        });
        allAgents.forEach((a: any) => {
          if (a.created_by == null && creatorById[String(a.id)] !== undefined)
            a.created_by = creatorById[String(a.id)];
        });
      }
    } catch {
      // created_by column not added yet — old rows stay editable.
    }
    const allowedAgents = isFullAccess
      ? allAgents
      : allAgents.filter((a: any) => {
          const matchUpa =
            uUpazilas.length === 0 || uUpazilas.includes(a.upazila);
          const matchDist =
            uDistricts.length === 0 || uDistricts.includes(a.district);
          return matchUpa && matchDist;
        });

    // Linked companies (agent_companies) — falls back to the single
    // company_id when the table does not exist yet.
    let agentCompaniesByAgent: Record<string, any[]> = {};
    try {
      const { data: acData, error: acErr } = await supabase
        .from("agent_companies")
        .select("*, companies(company_name, logo_url)")
        .order("created_at");
      if (!acErr && acData) {
        (acData as any[]).forEach((row: any) => {
          const key = String(row.agent_id);
          if (!agentCompaniesByAgent[key]) agentCompaniesByAgent[key] = [];
          agentCompaniesByAgent[key].push({
            ...row,
            company_id: String(row.company_id),
            company_name:
              row.companies?.company_name ||
              companyData?.find(
                (c: any) => String(c.id) === String(row.company_id)
              )?.company_name ||
              "",
            logo_url:
              row.companies?.logo_url ||
              companyData?.find(
                (c: any) => String(c.id) === String(row.company_id)
              )?.logo_url ||
              "",
          });
        });
      }
    } catch {
      // agent_companies table not created yet — legacy fallback applies.
    }
    const agentsWithLinks = allowedAgents.map((a: any) => ({
      ...a,
      agent_companies: agentCompaniesByAgent[String(a.id)] || [],
    }));

    setAgents(agentsWithLinks);
    setBazars(bazarData || []);
    setCompanies(companyData || []);

    const allDistricts = districtData || [];
    const allowedDists = isFullAccess
      ? allDistricts
      : allDistricts.filter(
          (d: any) => uDistricts.length === 0 || uDistricts.includes(d.name)
        );

    setDistricts(allowedDists);

    if (!districtFilter) {
      if (allowedDists.length === 1) {
        setDistrictFilter(allowedDists[0].name);
      } else if (
        allowedDists.some((d: any) => d.name === "Mymensingh") &&
        isFullAccess
      ) {
        setDistrictFilter("Mymensingh");
      }
    }
  }

  useEffect(() => {
    loadData();
  }, [role]);

  useEffect(() => {
    const total =
      Number(form.broiler_sale || 0) +
      Number(form.layer_sale || 0) +
      Number(form.sonali_sale || 0) +
      Number(form.cattle_sale || 0) +
      Number(form.fish_sale || 0);

    setForm((prev) => ({
      ...prev,
      monthly_sale: total.toString(),
    }));
  }, [
    form.broiler_sale,
    form.layer_sale,
    form.sonali_sale,
    form.cattle_sale,
    form.fish_sale,
  ]);

  function clearFilters() {
    setCompanyFilter("");
    setDistrictFilter(districts.length === 1 ? districts[0].name : "");
    setUpazilaFilter(
      !isFullAccess && assignedUpazilas.length === 1 ? assignedUpazilas[0] : ""
    );
    setUnionFilter("");
    setBazarFilter("");
    setExclusivityFilter("");
    setSearch("");
  }

  // ==================================================
  // AGENT TRACKING NOTES
  // ==================================================
  async function loadAgentNotes(agentId: number) {
    setActiveTrackingAgentId(agentId);
    setLoadingNotes(true);
    try {
      const { data, error } = await supabase
        .from("agent_trackings")
        .select("*")
        .eq("agent_id", agentId)
        .order("created_at", { ascending: true });

      if (!error) {
        setAgentNotes(data || []);
      } else {
        setAgentNotes([]);
      }
    } catch {
      setAgentNotes([]);
    } finally {
      setLoadingNotes(false);
    }
  }

  useEffect(() => {
    if (selectedAgent?.id) {
      setNewNoteText("");
      loadAgentNotes(selectedAgent.id);
    }
  }, [selectedAgent]);

  async function handleAddAgentNote(
    e: React.FormEvent,
    targetAgentId?: number
  ) {
    e.preventDefault();
    const aId = targetAgentId || selectedAgent?.id || activeTrackingAgentId;
    if (!aId || !newNoteText.trim()) return;

    setSavingNote(true);
    try {
      const { error } = await supabase.from("agent_trackings").insert([
        {
          agent_id: aId,
          note: newNoteText.trim(),
          user_id: currentUserId || null,
          user_name: currentUserName,
        },
      ]);

      if (error) throw error;
      setNewNoteText("");
      await loadAgentNotes(aId);
    } catch (err: any) {
      alert(err?.message || "Failed to save tracking note.");
    } finally {
      setSavingNote(false);
    }
  }

  // ✅ Map বাটনে ক্লিক করলে শুধু ওই এজেন্টে জুম করবে (নিজে থেকে পপআপ ওপেন হবে না)
  function handleViewAgentOnMap(agent: any) {
    if (!agent.latitude || !agent.longitude) {
      alert("Location not available for this agent.");
      return;
    }
    const lat = Number(agent.latitude);
    const lng = Number(agent.longitude);
    setSelectedAgent(null);
    setShowSummary(false);
    setShowMap(true);
    setMyLocationMode(false);
    setSelectedLocation([lat, lng]);
    setSelectedMapAgentId(agent.id);
    loadAgentNotes(agent.id);

    // আগে খোলা থাকা যেকোনো পপআপ বন্ধ করে দেওয়া হচ্ছে যাতে ম্যাপ পরিষ্কার দেখা যায়
    Object.values(agentMarkerRefs.current).forEach((m) => (m as any)?.closePopup());

    setTimeout(() => {
      mapSectionRef.current?.scrollIntoView({
        behavior: "smooth",
        block: "center",
      });
    }, 300);
  }

  // ==================================================
  // DEPENDENT LOCATION DROPDOWNS
  // ==================================================
  const selectedDistrict = districts.find((d) => d.name === form.district);

  useEffect(() => {
    if (!form.district) {
      setUpazilas([]);
      return;
    }
    async function loadUpazilas() {
      const district = districts.find((d) => d.name === form.district);
      if (!district) return;

      const { data } = await supabase
        .from("upazilas")
        .select("*")
        .eq("district_id", district.id)
        .order("name");

      const list = data || [];
      const allowed = isFullAccess
        ? list
        : list.filter(
            (u: any) =>
              assignedUpazilas.length === 0 || assignedUpazilas.includes(u.name)
          );
      setUpazilas(allowed);
    }
    loadUpazilas();
  }, [form.district, districts, isFullAccess, assignedUpazilas]);

  const selectedUpazila = upazilas.find((u: any) => u.name === form.upazila);

  useEffect(() => {
    if (!form.upazila) {
      setUnions([]);
      return;
    }
    async function loadUnions() {
      const upazila = upazilas.find((u) => u.name === form.upazila);
      if (!upazila) return;

      const { data } = await supabase
        .from("unions")
        .select("*")
        .eq("upazila_id", upazila.id)
        .order("name");

      setUnions(data || []);
    }
    loadUnions();
  }, [form.upazila, upazilas]);

  const selectedUnion = unions.find((u: any) => u.name === form.union_name);
  const filteredBazars = bazars.filter(
    (b: any) => b.union_id === selectedUnion?.id
  );

  useEffect(() => {
    async function loadFilterUpazilas() {
      let query = supabase.from("upazilas").select("*").order("name");

      if (districtFilter) {
        const district = districts.find((d) => d.name === districtFilter);
        if (!district) {
          setFilterUpazilas([]);
          return;
        }
        query = query.eq("district_id", district.id);
      } else if (!isFullAccess && districts.length > 0) {
        const distIds = districts.map((d) => d.id);
        query = query.in("district_id", distIds);
      }

      const { data } = await query;
      const list = data || [];
      const allowed = isFullAccess
        ? list
        : list.filter(
            (u: any) =>
              assignedUpazilas.length === 0 || assignedUpazilas.includes(u.name)
          );
      setFilterUpazilas(allowed);
    }
    loadFilterUpazilas();
  }, [districtFilter, districts, isFullAccess, assignedUpazilas]);

  useEffect(() => {
    async function loadFilterUnions() {
      if (upazilaFilter) {
        const upazila = filterUpazilas.find((u) => u.name === upazilaFilter);
        if (!upazila) {
          setFilterUnions([]);
          return;
        }
        const { data } = await supabase
          .from("unions")
          .select("*")
          .eq("upazila_id", upazila.id)
          .order("name");
        setFilterUnions(data || []);
        return;
      }

      const upazilaIds = filterUpazilas.map((u: any) => u.id);
      if (upazilaIds.length === 0) {
        setFilterUnions([]);
        return;
      }

      const { data } = await supabase
        .from("unions")
        .select("*")
        .in("upazila_id", upazilaIds)
        .order("name");
      setFilterUnions(data || []);
    }
    loadFilterUnions();
  }, [upazilaFilter, filterUpazilas]);

  useEffect(() => {
    if (!unionFilter) {
      setFilterBazars([]);
      return;
    }
    async function loadFilterBazars() {
      const union = filterUnions.find((u) => u.name === unionFilter);
      if (!union) return;

      const { data } = await supabase
        .from("bazars")
        .select("*")
        .eq("union_id", union.id)
        .order("name");

      setFilterBazars(data || []);
    }
    loadFilterBazars();
  }, [unionFilter, filterUnions]);

  // ---------- Linked companies in the form (exclusive / non-exclusive) ----------
  function syncFormFromCompanies(list: any[]) {
    const total = list.reduce((s: number, c: any) => s + linkSaleTotal(c), 0);
    const sums: Record<string, string> = {};
    SALE_KEYS.forEach((k) => {
      sums[k] = String(
        list.reduce((s: number, c: any) => s + Number(c[k] || 0), 0)
      );
    });
    setForm((prev) => ({
      ...prev,
      company_id: list[0]?.company_id ? String(list[0].company_id) : "",
      ...sums,
      monthly_sale: String(total),
    }));
  }

  function updateFormCompany(idx: number, patch: any) {
    const next = formCompanies.map((c: any, i: number) =>
      i === idx ? { ...c, ...patch } : c
    );
    if (next[idx]) {
      next[idx] = { ...next[idx], monthly_sale: String(linkSaleTotal(next[idx])) };
    }
    setFormCompanies(next);
    syncFormFromCompanies(next);
  }

  function addFormCompany(companyId: string) {
    if (!companyId) return;
    if (
      formCompanies.some(
        (c: any) => String(c.company_id) === String(companyId)
      )
    ) {
      alert("Company already added.");
      return;
    }
    const picked = companies.find((c) => String(c.id) === String(companyId));
    const next = [
      ...formCompanies,
      {
        ...emptyCompanyLink(companyId),
        company_name: picked?.company_name || "",
        logo_url: picked?.logo_url || "",
      },
    ];
    setFormCompanies(next);
    syncFormFromCompanies(next);
  }

  function removeFormCompany(idx: number) {
    if (formCompanies.length <= 1) return;
    const next = formCompanies.filter((_: any, i: number) => i !== idx);
    setFormCompanies(next);
    syncFormFromCompanies(next);
  }

  function handlePrimaryCompanyChange(companyId: string) {
    if (formCompanies.length === 0) {
      const next = [{ ...emptyCompanyLink(companyId) }];
      setFormCompanies(next);
      syncFormFromCompanies(next);
    } else {
      updateFormCompany(0, { company_id: companyId });
    }
  }

  function handleTypeChange(t: "exclusive" | "nonexclusive") {
    setFormType(t);
    if (t === "exclusive" && formCompanies.length > 1) {
      const next = [formCompanies[0]];
      setFormCompanies(next);
      syncFormFromCompanies(next);
    }
  }

  // ==================================================
  // SAVE, DELETE & LOCATE ME (INCLUDING NOURISH AGENT ID)
  // ==================================================
  async function handleSave() {
    if (savingAgent) return;
    setSavingAgent(true);
    try {
      // Permission guard (buttons are hidden too — this is the
      // backstop for the save itself). Agent add is NOT
      // category-gated (rule 5, like farmers).
      if (editingId) {
        const editingAgent = agents.find(
          (a: any) => Number(a.id) === Number(editingId)
        );
        if (!canEditAgent(editingAgent || {})) {
          alert("You can only edit agents you added.");
          return;
        }
      } else if (!canAddAgentHere) {
        alert("You do not have permission to add agents.");
        return;
      }
      const links =
        formCompanies.length > 0
          ? formCompanies.filter((c: any) => c.company_id)
          : form.company_id
          ? [
              {
                ...emptyCompanyLink(form.company_id),
                broiler_sale: form.broiler_sale,
                sonali_sale: form.sonali_sale,
                layer_sale: form.layer_sale,
                cattle_sale: form.cattle_sale,
                fish_sale: form.fish_sale,
                monthly_sale: form.monthly_sale,
              },
            ]
          : [];
      if (
        links.length === 0 ||
        !form.agent_name.trim() ||
        !form.district ||
        !form.upazila
      ) {
        alert("Company, Agent Name, District and Upazila are required.");
        return;
      }

      const sums: Record<string, number> = {};
      SALE_KEYS.forEach((k) => {
        sums[k] = links.reduce((s: number, c: any) => s + Number(c[k] || 0), 0);
      });
      const totalSale = links.reduce(
        (s: number, c: any) => s + linkSaleTotal(c),
        0
      );
      const primaryCompanyId = String(links[0].company_id);
      const primaryCompanyObj = companies.find(
        (c) => String(c.id) === primaryCompanyId
      );
      const cleanAgentCode = isNourishName(primaryCompanyObj?.company_name)
        ? form.agent_code.trim()
        : "";

      const payload: any = {
        ...form,
        agent_code: cleanAgentCode,
        company_id: Number(primaryCompanyId),
        broiler_sale: sums.broiler_sale,
        layer_sale: sums.layer_sale,
        sonali_sale: sums.sonali_sale,
        cattle_sale: sums.cattle_sale,
        fish_sale: sums.fish_sale,
        monthly_sale: totalSale,
        latitude: form.latitude === "" ? null : Number(form.latitude),
        longitude: form.longitude === "" ? null : Number(form.longitude),
      };

      let agentId: number | null = editingId;
      if (editingId) {
        await updateAgent(editingId, payload);
        // Ownership (rule 2): stamp created_by on first edit of an
        // old row; never overwrite an existing creator. Best-effort:
        // a missing created_by column must not break the save.
        let stampCreatedBy: Record<string, any> = {};
        try {
          const { data: existing } = await supabase
            .from("agents")
            .select("created_by")
            .eq("id", editingId)
            .maybeSingle();
          if (existing && !existing.created_by && currentUserId) {
            stampCreatedBy = { created_by: currentUserId };
          }
        } catch {
          // created_by column not added yet.
        }
        await supabase
          .from("agents")
          .update({ agent_code: cleanAgentCode, ...stampCreatedBy })
          .eq("id", editingId);
      } else {
        const created: any = await createAgent(payload);
        const createdId =
          created?.id || created?.data?.id || created?.[0]?.id;
        if (createdId) {
          agentId = Number(createdId);
          await supabase
            .from("agents")
            .update({ agent_code: cleanAgentCode })
            .eq("id", createdId);
        } else {
          const { data: found } = await supabase
            .from("agents")
            .select("id")
            .eq("agent_name", form.agent_name.trim())
            .order("created_at", { ascending: false })
            .limit(1);
          if (found && found.length > 0) agentId = Number(found[0].id);
          if (cleanAgentCode && agentId) {
            await supabase
              .from("agents")
              .update({ agent_code: cleanAgentCode })
              .eq("id", agentId);
          }
        }
      }

      // Ownership: new agents are stamped with their creator.
      if (!editingId && agentId && currentUserId) {
        try {
          await supabase
            .from("agents")
            .update({ created_by: currentUserId })
            .eq("id", agentId);
        } catch {
          // created_by column not added yet.
        }
      }

      // Save per-company links (agent_companies). The agent row above
      // already holds the summed sales, so the agent is safe even if
      // this table does not exist yet.
      // Company links inherit the agent's creator (no separate gate).
      const agentCreatorId = editingId
        ? agents.find((a: any) => Number(a.id) === Number(editingId))
            ?.created_by || currentUserId
        : currentUserId;
      if (agentId) {
        try {
          await supabase
            .from("agent_companies")
            .delete()
            .eq("agent_id", agentId);
          const rows = links.map((c: any) => ({
            agent_id: agentId,
            company_id: Number(c.company_id),
            broiler_sale: Number(c.broiler_sale || 0),
            sonali_sale: Number(c.sonali_sale || 0),
            layer_sale: Number(c.layer_sale || 0),
            cattle_sale: Number(c.cattle_sale || 0),
            fish_sale: Number(c.fish_sale || 0),
            monthly_sale: linkSaleTotal(c),
            created_by: agentCreatorId || null,
          }));
          let { error: linkErr } = await supabase
            .from("agent_companies")
            .insert(rows);
          if (linkErr && /created_by/i.test(linkErr.message || "")) {
            const rowsWithoutCreator = rows.map((r: any) => {
              const { created_by, ...rest } = r;
              return rest;
            });
            const retry = await supabase
              .from("agent_companies")
              .insert(rowsWithoutCreator);
            linkErr = retry.error;
          }
          if (linkErr) throw linkErr;
        } catch (linkErr: any) {
          alert(
            "Agent saved, but company links could not be saved. Please run agent_companies-migration.sql in Supabase first, then edit the agent again to add companies."
          );
        }
      }

      alert(editingId ? "Agent Updated" : "Agent Saved");
      setShowForm(false);
      setEditingId(null);
      setForm({ ...EMPTY_FORM });
      setFormCompanies([]);
      setFormType("exclusive");
      loadData();
    } catch (error: any) {
      alert(error?.message || "Save Failed");
    } finally {
      setSavingAgent(false);
    }
  }

  async function handleDelete(id: number) {
    const target = agents.find((a: any) => Number(a.id) === Number(id));
    if (!canEditAgent(target || {})) {
      alert("You can only delete agents you added.");
      return;
    }
    const { data: linkedFarmers } = await supabase
      .from("farmers")
      .select("id")
      .eq("agent_id", id)
      .limit(1);
    if (linkedFarmers && linkedFarmers.length > 0) {
      alert("This agent has farmers.");
      return;
    }
    if (!confirm("Delete Agent?")) return;
    await deleteAgent(id);
    setSelectedAgent(null);
    loadData();
  }

  // Universal Location Helper (Works on both Android Native App & Browser)
  async function getDeviceCoordinates(): Promise<{ lat: number; lng: number }> {
    if (Capacitor.isNativePlatform()) {
      let perm = await Geolocation.checkPermissions();
      if (perm.location !== "granted") {
        perm = await Geolocation.requestPermissions();
      }
      if (perm.location !== "granted") {
        throw new Error(
          "Location permission denied. Please allow GPS access."
        );
      }

      try {
        const pos = await Geolocation.getCurrentPosition({
          enableHighAccuracy: true,
          timeout: 10000,
          maximumAge: 0,
        });
        return { lat: pos.coords.latitude, lng: pos.coords.longitude };
      } catch {
        const fallbackPos = await Geolocation.getCurrentPosition({
          enableHighAccuracy: false,
          timeout: 10000,
        });
        return {
          lat: fallbackPos.coords.latitude,
          lng: fallbackPos.coords.longitude,
        };
      }
    } else {
      return new Promise((resolve, reject) => {
        if (!navigator.geolocation) {
          reject(new Error("Geolocation is not supported by your browser."));
          return;
        }
        navigator.geolocation.getCurrentPosition(
          (pos) =>
            resolve({ lat: pos.coords.latitude, lng: pos.coords.longitude }),
          (err) => reject(err),
          { enableHighAccuracy: true, timeout: 10000 }
        );
      });
    }
  }

  async function handleLocateMeInPicker() {
    setLocatingInPicker(true);
    try {
      const { lat, lng } = await getDeviceCoordinates();
      setPickerCenterOverride([lat, lng]);
      setForm((prev) => ({
        ...prev,
        latitude: lat.toFixed(6),
        longitude: lng.toFixed(6),
      }));
    } catch (err: any) {
      alert(
        err?.message ||
          "Unable to retrieve your location. Please allow GPS access."
      );
    } finally {
      setLocatingInPicker(false);
    }
  }

  // ✅ Locate Me ক্লিক করলে selectedLocation মুছে ফেলবে না, ফলে সিলেক্ট করা এজেন্টের সাথে ডিসট্যান্স ও Polyline দেখাবে
  async function getMyLocation() {
    try {
      const { lat, lng } = await getDeviceCoordinates();
      const coords: [number, number] = [lat, lng];
      setUserLocation(coords);
      setMyLocationMode(true);

      const baseAgents =
        selectedMapAgentId != null
          ? filteredAgents.filter(
              (a) => String(a.id) === String(selectedMapAgentId)
            )
          : filteredAgents;
      const nearest = baseAgents
        .filter((a) => a.latitude && a.longitude)
        .map((a) => {
          const km = calcDistanceKm(
            coords[0],
            coords[1],
            Number(a.latitude),
            Number(a.longitude)
          );
          return {
            ...a,
            distanceKm: km,
            distanceText: formatDistance(km),
          };
        })
        .sort((a, b) => a.distanceKm - b.distanceKm)
        .slice(0, 10);

      setNearestAgents(nearest);
    } catch (err: any) {
      alert(err?.message || "Location access denied");
    }
  }

  // ==================================================
  // FILTERED AGENTS & COMPANY-WISE GROUPING
  // ==================================================
  const filteredAgents = agents.filter((agent) => {
    const q = search.toLowerCase();
    const agentLinks = getAgentCompanies(agent);
    const matchSearch =
      !q ||
      agent.agent_name?.toLowerCase().includes(q) ||
      agent.agent_code?.toLowerCase().includes(q) ||
      agent.proprietor?.toLowerCase().includes(q) ||
      agent.mobile?.toLowerCase().includes(q) ||
      agent.bazar?.toLowerCase().includes(q) ||
      agentLinks.some((c: any) =>
        (c.company_name || "").toLowerCase().includes(q)
      );

    const matchCompany =
      companyFilter === "" ||
      agentLinks.some((c: any) => String(c.company_id) === companyFilter);
    const matchExclusivity =
      exclusivityFilter === "" ||
      (exclusivityFilter === "nonexclusive"
        ? agentLinks.length >= 2
        : agentLinks.length < 2);
    const matchDistrict =
      districtFilter === "" || agent.district === districtFilter;
    const matchUpazila =
      upazilaFilter === "" || agent.upazila === upazilaFilter;
    const matchUnion =
      unionFilter === "" || agent.union_name === unionFilter;
    const matchBazar = bazarFilter === "" || agent.bazar === bazarFilter;

    return (
      matchSearch &&
      matchCompany &&
      matchExclusivity &&
      matchDistrict &&
      matchUpazila &&
      matchUnion &&
      matchBazar
    );
  });

  const activeDistanceText =
    userLocation && selectedLocation
      ? formatDistance(
          calcDistanceKm(
            userLocation[0],
            userLocation[1],
            selectedLocation[0],
            selectedLocation[1]
          )
        )
      : null;

  // Locate Me + a selected agent: show ONLY that agent on the map
  // (all other agent markers hidden) until the selection is cleared.
  const isolateSelectedAgent = Boolean(
    myLocationMode &&
      userLocation &&
      selectedLocation &&
      selectedMapAgentId != null
  );
  const mapAgents = isolateSelectedAgent
    ? filteredAgents.filter(
        (a: any) => String(a.id) === String(selectedMapAgentId)
      )
    : filteredAgents;

  const sortedAgents = [...filteredAgents].sort((a: any, b: any) => {
    const companyCompare = (
      a.companies?.company_name || ""
    ).localeCompare(b.companies?.company_name || "");
    if (companyCompare !== 0) return companyCompare;
    return (a.agent_name || "").localeCompare(b.agent_name || "");
  });

  // Grouped per agent-company LINK, so a non-exclusive agent appears
  // under every company it is linked to (not only its primary one).
  const companyGroupedAgents = Object.values(
    sortedAgents.reduce((acc: Record<string, any>, agent: any) => {
      getAgentCompanies(agent).forEach((link: any) => {
        const key = String(link.company_id || "unassigned");
        if (!acc[key]) {
          const compName = link.company_name || "Unassigned Company";
          acc[key] = {
            companyId: key,
            companyName: compName,
            isNourish: isNourishName(compName),
            logoUrl: link.logo_url || "",
            totalSale: 0,
            agents: [],
          };
        }
        acc[key].agents.push({ agent, link });
        acc[key].totalSale += Number(
          (link.monthly_sale ?? linkSaleTotal(link)) || 0
        );
      });
      return acc;
    }, {})
  );

  const selectedCompanyName =
    companies.find((c) => String(c.id) === companyFilter)?.company_name ||
    "All Companies";

  const nonExclusiveCount = filteredAgents.filter((a: any) =>
    isNonExclusive(a)
  ).length;

  const nonExPopupAgent =
    agents.find((a: any) => a.id === nonExPopupAgentId) || null;

  const totalMonthlySale = filteredAgents.reduce(
    (sum, agent) => sum + Number(agent.monthly_sale || 0),
    0
  );

  // Summary per agent-company LINK (no double-counting: each link's
  // own sales count once under its company; agents counted once per
  // company via a Set of distinct agent ids).
  const companySummary = Object.values(
    filteredAgents.reduce((acc: any, agent) => {
      getAgentCompanies(agent).forEach((link: any) => {
        const company = link.company_name || "Unknown";
        if (!acc[company]) {
          acc[company] = {
            name: company,
            agentIds: new Set(),
            agents: 0,
            broiler: 0,
            layer: 0,
            sonali: 0,
            cattle: 0,
            fish: 0,
            sale: 0,
          };
        }
        acc[company].agentIds.add(String(agent.id));
        acc[company].agents = acc[company].agentIds.size;
        acc[company].broiler += Number(link.broiler_sale || 0);
        acc[company].layer += Number(link.layer_sale || 0);
        acc[company].sonali += Number(link.sonali_sale || 0);
        acc[company].cattle += Number(link.cattle_sale || 0);
        acc[company].fish += Number(link.fish_sale || 0);
        acc[company].sale += Number(
          (link.monthly_sale ?? linkSaleTotal(link)) || 0
        );
      });
      return acc;
    }, {})
  );

  // ==================================================
  // EXCEL EXPORT
  // ==================================================
  async function exportAgentsExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Agents");

      worksheet.getRow(1).values = [
        "Company",
        "Agent ID",
        "Agent Name",
        "Proprietor",
        "Mobile",
        "Bazar",
        "District",
        "Upazila",
        "Union",
        "Broiler Sale",
        "Layer Sale",
        "Sonali Sale",
        "Cattle Sale",
        "Fish Sale",
        "Total Monthly Sale",
        "Focus Item",
        "Status",
        "Companies",
        "Type",
      ];

      sortedAgents.forEach((agent: any) => {
        worksheet.addRow([
          agent.companies?.company_name || "",
          agent.agent_code || "-",
          agent.agent_name || "",
          agent.proprietor || "",
          agent.mobile || "",
          agent.bazar || "",
          agent.district || "",
          agent.upazila || "",
          agent.union_name || "",
          Number(agent.broiler_sale || 0),
          Number(agent.layer_sale || 0),
          Number(agent.sonali_sale || 0),
          Number(agent.cattle_sale || 0),
          Number(agent.fish_sale || 0),
          Number(agent.monthly_sale || 0),
          agent.focus_item || "",
          agent.status || "",
          getAgentCompanies(agent)
            .map((c: any) => c.company_name)
            .filter(Boolean)
            .join(", "),
          isNonExclusive(agent) ? "Non-exclusive" : "Exclusive",
        ]);
      });

      for (let col = 1; col <= 19; col++) {
        const cell = worksheet.getCell(1, col);
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: {
            argb:
              col === 15
                ? "15803D"
                : col >= 10 && col <= 14
                ? "D97706"
                : "1E3A8A",
          },
        };
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.alignment = {
          horizontal: "center",
          vertical: "middle",
          wrapText: true,
        };
        cell.border = {
          top: { style: "thin", color: { argb: "FFFFFF" } },
          left: { style: "thin", color: { argb: "FFFFFF" } },
          bottom: { style: "thin", color: { argb: "FFFFFF" } },
          right: { style: "thin", color: { argb: "FFFFFF" } },
        };
      }

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber >= 2) {
          row.eachCell((cell) => {
            cell.border = {
              top: { style: "thin", color: { argb: "D1D5DB" } },
              left: { style: "thin", color: { argb: "D1D5DB" } },
              bottom: { style: "thin", color: { argb: "D1D5DB" } },
              right: { style: "thin", color: { argb: "D1D5DB" } },
            };
            cell.alignment = { vertical: "middle", wrapText: true };
          });
        }
      });

      const widths = [
        24, 14, 22, 20, 16, 20, 18, 18, 18, 15, 15, 15, 15, 15, 20, 20, 14,
        26, 14,
      ];
      widths.forEach((w, idx) => (worksheet.getColumn(idx + 1).width = w));
      worksheet.getRow(1).height = 30;
      worksheet.views = [{ state: "frozen", ySplit: 1 }];
      worksheet.autoFilter = { from: "A1", to: "S1" };

      // Second sheet: one row per agent-company link.
      const linkSheet = workbook.addWorksheet("Agent Companies");
      linkSheet.getRow(1).values = [
        "Agent",
        "Company",
        "B",
        "S",
        "L",
        "C",
        "F",
        "Total",
      ];
      sortedAgents.forEach((agent: any) => {
        getAgentCompanies(agent).forEach((link: any) => {
          linkSheet.addRow([
            agent.agent_name || "",
            link.company_name || "",
            Number(link.broiler_sale || 0),
            Number(link.sonali_sale || 0),
            Number(link.layer_sale || 0),
            Number(link.cattle_sale || 0),
            Number(link.fish_sale || 0),
            Number((link.monthly_sale ?? linkSaleTotal(link)) || 0),
          ]);
        });
      });
      linkSheet.getRow(1).eachCell((cell) => {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "1E3A8A" },
        };
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
      });
      [22, 24, 10, 10, 10, 10, 10, 14].forEach((w, idx) =>
        linkSheet.getColumn(idx + 1).width = w
      );
      linkSheet.views = [{ state: "frozen", ySplit: 1 }];
      linkSheet.autoFilter = { from: "A1", to: "H1" };

      const today = new Date().toISOString().split("T")[0];
      const safeCompanyName = (selectedCompanyName || "Report").replace(
        /[\s\/\\:@?*]+/g,
        "_"
      );
      const fileName = `Agents_${safeCompanyName}_${today}.xlsx`;

      const excelBuffer = await workbook.xlsx.writeBuffer();

      if (Capacitor.isNativePlatform()) {
        let binary = "";
        const bytes = new Uint8Array(excelBuffer);
        const chunkSize = 0x8000;
        for (let i = 0; i < bytes.length; i += chunkSize) {
          binary += String.fromCharCode.apply(
            null,
            Array.from(bytes.subarray(i, i + chunkSize))
          );
        }
        const base64Data = window.btoa(binary);

        const savedFile = await Filesystem.writeFile({
          path: fileName,
          data: base64Data,
          directory: Directory.Documents,
          recursive: true,
        });

        alert(`File saved successfully:\n📁 Documents/${fileName}`);

        if (savedFile.uri) {
          try {
            await Share.share({
              title: fileName,
              url: savedFile.uri,
              dialogTitle: "Open or View Excel File",
            });
          } catch {}
        }
      } else {
        const blob = new Blob([excelBuffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        saveAs(blob, fileName);
      }
    } catch (error) {
      console.error("Agent Excel Export Error:", error);
      alert("Agent Excel export failed.");
    }
  }

  function openEditForm(agent: any) {
    if (!canEditAgent(agent)) {
      alert("You can only edit agents you added.");
      return;
    }
    const links = getAgentCompanies(agent);
    const formLinks = links.map((c: any) => ({
      id: String(c.id || "").startsWith("legacy-") ? undefined : c.id,
      company_id: String(c.company_id),
      company_name: c.company_name || "",
      logo_url: c.logo_url || "",
      broiler_sale: String(c.broiler_sale ?? 0),
      sonali_sale: String(c.sonali_sale ?? 0),
      layer_sale: String(c.layer_sale ?? 0),
      cattle_sale: String(c.cattle_sale ?? 0),
      fish_sale: String(c.fish_sale ?? 0),
      monthly_sale: String(c.monthly_sale ?? linkSaleTotal(c)),
    }));
    setEditingId(agent.id);
    setFormType(formLinks.length >= 2 ? "nonexclusive" : "exclusive");
    setFormCompanies(formLinks);
    setForm({
      company_id: String(agent.company_id),
      agent_code: agent.agent_code || "",
      agent_name: agent.agent_name || "",
      proprietor: agent.proprietor || "",
      mobile: agent.mobile || "",
      bazar: agent.bazar || "",
      union_name: agent.union_name || "",
      upazila: agent.upazila || "",
      district: agent.district || "",
      broiler_sale: String(agent.broiler_sale || 0),
      layer_sale: String(agent.layer_sale || 0),
      sonali_sale: String(agent.sonali_sale || 0),
      cattle_sale: String(agent.cattle_sale || 0),
      fish_sale: String(agent.fish_sale || 0),
      monthly_sale: String(agent.monthly_sale || 0),
      focus_item: agent.focus_item || "",
      latitude: String(agent.latitude || ""),
      longitude: String(agent.longitude || ""),
      status: agent.status || "Active",
    });
    setShowForm(true);
  }

  return (
    <div className="p-2.5 sm:p-4 max-w-7xl mx-auto space-y-3">
      {/* ==================================================
          1. COMPACT FILTER BAR AT THE TOP
      ================================================== */}
      <div className="bg-white border border-slate-300 rounded-xl p-2 shadow-sm">
        <button
          type="button"
          onClick={() => setShowFilters(!showFilters)}
          className="sm:hidden inline-flex items-center gap-1 rounded-lg border border-slate-300 bg-slate-50 px-2.5 py-1 text-[11px] font-semibold text-slate-700"
        >
          <Search size={12} /> Filters {showFilters ? "▲" : "▼"}
        </button>
        <div
          className={`${showFilters ? "grid" : "hidden"} sm:grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-1.5 items-center mt-2 sm:mt-0`}
        >
          <div className="col-span-2 sm:col-span-2 lg:col-span-2 relative">
            <Search
              size={13}
              className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400"
            />
            <input
              type="text"
              placeholder="Search agent name, ID, proprietor, mobile..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full rounded-lg border border-slate-300 pl-8 pr-2.5 py-1 text-xs outline-none focus:border-blue-500"
            />
          </div>

          <SearchableSelect
            searchable={false}
            value={districtFilter}
            onChange={(v) => {
              setDistrictFilter(v);
              setUpazilaFilter("");
              setUnionFilter("");
              setBazarFilter("");
            }}
            placeholder="All Districts"
            options={districts.map((d) => ({ value: d.name, label: d.name }))}
          />

          <SearchableSelect
            value={upazilaFilter}
            onChange={(v) => {
              setUpazilaFilter(v);
              setUnionFilter("");
              setBazarFilter("");
            }}
            placeholder="All Upazilas"
            options={filterUpazilas.map((u: any) => ({
              value: u.name,
              label: u.name,
            }))}
          />

          <SearchableSelect
            value={unionFilter}
            onChange={(v) => {
              setUnionFilter(v);
              setBazarFilter("");
            }}
            placeholder="All Unions"
            options={filterUnions.map((u: any) => ({
              value: u.name,
              label: u.name,
            }))}
          />

          <SearchableSelect
            value={bazarFilter}
            onChange={setBazarFilter}
            placeholder="All Bazars"
            options={filterBazars.map((b: any) => ({
              value: b.name,
              label: b.name,
            }))}
          />

          <SearchableSelect
            searchable={false}
            value={exclusivityFilter}
            onChange={setExclusivityFilter}
            placeholder="All Agents"
            options={[
              { value: "exclusive", label: "Exclusive only" },
              { value: "nonexclusive", label: "Non-exclusive only" },
            ]}
          />

          <div className="col-span-2 sm:col-span-1 flex items-center gap-1.5">
            <SearchableSelect
              className="w-full"
              value={companyFilter}
              onChange={setCompanyFilter}
              placeholder="All Companies"
              options={companies.map((c) => ({
                value: String(c.id),
                label: c.company_name,
              }))}
            />

            <button
              onClick={clearFilters}
              title="Reset Filters"
              className="shrink-0 inline-flex items-center justify-center rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-2.5 py-1 text-xs font-semibold text-slate-700 transition"
            >
              <RotateCcw size={13} />
            </button>
          </div>
        </div>
      </div>

      {/* ==================================================
          2. ADD AGENT, SUMMARY, EXPORT, VIEW MAP BAR
      ================================================== */}
      <div className="flex flex-wrap items-center justify-between gap-1.5 bg-white border border-slate-300 rounded-xl px-2.5 py-1.5 shadow-sm">
        <div className="flex flex-wrap items-center gap-1.5">
          {canAddAgentHere && (
          <button
            onClick={() => {
              setEditingId(null);
              const defDist =
                districts.length === 1
                  ? districts[0].name
                  : districts.some((d) => d.name === "Mymensingh")
                  ? "Mymensingh"
                  : "";
              const defUpa =
                !isFullAccess && assignedUpazilas.length === 1
                  ? assignedUpazilas[0]
                  : "";
              setForm({
                ...EMPTY_FORM,
                district: defDist,
                upazila: defUpa,
              });
              setFormCompanies([{ ...emptyCompanyLink("") }]);
              setFormType("exclusive");
              setCompanyPickerId("");
              setShowForm(true);
            }}
            className="inline-flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-2.5 py-1 rounded-lg text-[11px] font-semibold shadow-sm transition"
          >
            <Plus size={14} /> Add Agent
          </button>
          )}

          <button
            onClick={() => {
              setShowMap(false);
              setShowSummary(true);
            }}
            className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded-lg text-[11px] font-semibold shadow-sm transition"
          >
            <BarChart3 size={14} /> Summary
          </button>

          <button
            onClick={exportAgentsExcel}
            className="inline-flex items-center gap-1.5 bg-emerald-700 hover:bg-emerald-800 text-white px-2.5 py-1 rounded-lg text-[11px] font-semibold shadow-sm transition"
          >
            <FileSpreadsheet size={14} /> Export Excel
          </button>

          <button
            onClick={() => {
              if (!showMap) {
                setMyLocationMode(false);
                setUserLocation(null);
                setSelectedLocation(null);
                setSelectedMapAgentId(null);
              }
              setShowSummary(false);
              setShowMap(!showMap);
            }}
            className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-[11px] font-semibold shadow-sm transition ${
              showMap
                ? "bg-red-600 hover:bg-red-700 text-white"
                : "bg-blue-600 hover:bg-blue-700 text-white"
            }`}
          >
            <MapIcon size={14} />
            {showMap ? "Close Map" : "View Map"}
          </button>
        </div>

        <div className="flex items-center gap-3 text-xs font-semibold text-slate-600">
          <span>
            Total Agents:{" "}
            <strong className="text-slate-900">{filteredAgents.length}</strong>
          </span>
          <span>
            Non-exclusive:{" "}
            <strong className="text-violet-700">{nonExclusiveCount}</strong>
          </span>
          <span>
            Total Sale:{" "}
            <strong className="text-emerald-700">
              {totalMonthlySale.toLocaleString()}
            </strong>
          </span>
        </div>
      </div>

      {/* ==================================================
          3. INTERACTIVE AGENT MAP VIEW
      ================================================== */}
      {showMap && (
        <div
          ref={mapSectionRef}
          id="agent-map-section"
          className="bg-white rounded-xl shadow border border-slate-300 p-2 relative z-10"
        >
          <h3 className="text-sm font-bold text-slate-800 mb-1">
            Agent Locations
          </h3>

          <div className="w-full h-[320px] sm:h-[420px] lg:h-[480px] overflow-hidden rounded-xl border border-slate-300 relative">
            <MapContainer
              center={mapCenter}
              zoom={8}
              style={{ height: "100%", width: "100%", zIndex: 1 }}
            >
              <FitBounds
                skip={Boolean(selectedLocation || userLocation)}
                locations={mapAgents
                  .filter((agent) => agent.latitude && agent.longitude)
                  .map((agent) => ({
                    lat: Number(agent.latitude),
                    lng: Number(agent.longitude),
                  }))}
              />
              <MapLayers />

              {/* অটো ফোকাস ও বাউন্ডস কন্ট্রোলার */}
              <MapFocusController
                agentPos={selectedLocation}
                myPos={userLocation}
                myLocationMode={myLocationMode}
              />

              {userLocation && (
                <Marker
                  position={userLocation}
                  ref={(ref) => {
                    myLocationMarkerRef.current = ref;
                  }}
                >
                  <Tooltip
                    permanent
                    direction="top"
                    offset={[0, -15]}
                    interactive={true}
                    eventHandlers={{
                      click: () => {
                        myLocationMarkerRef.current?.openPopup();
                      },
                    }}
                  >
                    <span className="cursor-pointer font-semibold">
                      📍 My Location
                    </span>
                  </Tooltip>
                  <Popup>
                    <div className="max-h-60 overflow-y-auto w-52 text-xs space-y-1.5">
                      <h3 className="font-bold text-slate-800">
                        📍 Nearest Agents
                      </h3>
                      {nearestAgents.length === 0 ? (
                        <div>No nearby agents found.</div>
                      ) : (
                        nearestAgents.map((a, index) => (
                          <div key={a.id} className="border-b pb-1">
                            <button
                              type="button"
                              onClick={() => {
                                const targetPos: [number, number] = [
                                  Number(a.latitude),
                                  Number(a.longitude),
                                ];
                                setMyLocationMode(false);
                                setSelectedLocation(targetPos);
                                setSelectedMapAgentId(a.id);
                                setNewNoteText("");
                                loadAgentNotes(a.id);
                                myLocationMarkerRef.current?.closePopup();
                              }}
                              className="text-blue-600 font-bold hover:underline text-left"
                            >
                              {index + 1}. {a.agent_name}
                            </button>
                            <div className="text-slate-500">
                              {a.distanceText} away
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </Popup>
                </Marker>
              )}

              {/* My Location থেকে Selected Agent পর্যন্ত ড্যাশ-ড্যাশ লাইন */}
              {userLocation && selectedLocation && (
                <Polyline
                  positions={[userLocation, selectedLocation]}
                  pathOptions={{
                    color: "#2563eb",
                    weight: 3,
                    dashArray: "8, 8",
                  }}
                />
              )}

              {mapAgents
                .filter((a) => a.latitude && a.longitude)
                .map((agent) => {
                  const aLat = Number(agent.latitude);
                  const aLng = Number(agent.longitude);

                  return (
                    <Marker
                      key={agent.id}
                      ref={(ref) => {
                        agentMarkerRefs.current[agent.id] = ref;
                      }}
                      position={[aLat, aLng]}
                      icon={getAgentMarkerIcon(agent)}
                      eventHandlers={{
                        click: () => {
                          setSelectedLocation([aLat, aLng]);
                          setSelectedMapAgentId(agent.id);
                          setNewNoteText("");
                          loadAgentNotes(agent.id);
                        },
                      }}
                    >
                      <Popup maxWidth={340}>
                        <div className="w-[270px] sm:w-[290px] text-xs space-y-2">
                          <div className="flex items-center gap-2 border-b pb-1.5">
                            {agent.companies?.logo_url && (
                              <img
                                src={agent.companies.logo_url}
                                alt="logo"
                                className="w-9 h-9 object-contain border rounded p-0.5 bg-white shrink-0"
                              />
                            )}
                            <div>
                              <h3 className="font-bold text-sm text-slate-800 leading-tight">
                                {agent.agent_name}
                                {agent.agent_code
                                  ? ` (ID: ${agent.agent_code})`
                                  : ""}
                              </h3>
                              <p className="text-[11px] text-slate-500">
                                {getAgentCompanies(agent)
                                  .map((c: any) => c.company_name)
                                  .filter(Boolean)
                                  .join(", ")}
                                {isNonExclusive(agent) && (
                                  <span className="ml-1 rounded bg-violet-100 px-1 py-0.5 text-[9px] font-bold italic text-violet-700">
                                    non-ex.
                                  </span>
                                )}
                                {agent.proprietor
                                  ? ` • Prop: ${agent.proprietor}`
                                  : ""}
                              </p>
                            </div>
                          </div>

                          {/* Agent Tracking Box */}
                          <div className="bg-emerald-50 border border-emerald-200 rounded-lg p-2 space-y-1.5">
                            <div className="font-bold text-[11px] text-emerald-900">
                              📝 Agent Tracking
                            </div>
                            <div className="max-h-24 overflow-y-auto space-y-1 pr-0.5">
                              {loadingNotes &&
                              activeTrackingAgentId === agent.id ? (
                                <div className="text-[11px] text-slate-400">
                                  Loading...
                                </div>
                              ) : activeTrackingAgentId === agent.id &&
                                agentNotes.length > 0 ? (
                                agentNotes.map((n) => (
                                  <div
                                    key={n.id}
                                    className="bg-white border border-emerald-100 rounded px-2 py-1 text-[11px] flex items-start gap-1.5"
                                  >
                                    <div className="mt-0.5 shrink-0">
                                      <UserAvatar
                                        name={n.user_name || "Officer"}
                                        avatarUrl={getUserAvatar(
                                          n.user_id,
                                          n.user_name
                                        )}
                                        size="w-4 h-4"
                                        textSize="text-[8px]"
                                      />
                                    </div>
                                    <div className="min-w-0 flex-1 break-words">
                                      <span className="font-bold text-emerald-800">
                                        {n.user_name}:
                                      </span>{" "}
                                      <span className="text-slate-700">
                                        {renderNoteWithMentions(n.note)}
                                      </span>
                                    </div>
                                  </div>
                                ))
                              ) : (
                                <div className="text-[10px] text-slate-400">
                                  No tracking note yet.
                                </div>
                              )}
                            </div>

                            <form
                              onSubmit={(e) => handleAddAgentNote(e, agent.id)}
                              className="flex gap-1"
                            >
                              <input
                                type="text"
                                placeholder="Write note..."
                                value={
                                  activeTrackingAgentId === agent.id
                                    ? newNoteText
                                    : ""
                                }
                                onChange={(e) => {
                                  setActiveTrackingAgentId(agent.id);
                                  setNewNoteText(e.target.value);
                                }}
                                className="flex-1 border border-emerald-300 rounded px-2 py-1 text-[11px] bg-white outline-none"
                              />
                              <button
                                type="submit"
                                disabled={savingNote}
                                className="bg-emerald-600 text-white px-2 py-1 rounded text-[11px] font-semibold cursor-pointer disabled:opacity-60"
                              >
                                {savingNote ? "Saving..." : "Track"}
                              </button>
                            </form>
                          </div>

                          <table className="w-full border-collapse border border-slate-200 text-[11px]">
                            <tbody>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Mobile
                                </td>
                                <td className="border px-2 py-1">
                                  {agent.mobile || "-"}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Address
                                </td>
                                <td className="border px-2 py-1">
                                  {[
                                    agent.bazar,
                                    agent.union_name,
                                    agent.upazila,
                                    agent.district,
                                  ]
                                    .filter(Boolean)
                                    .join(", ")}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Companies
                                </td>
                                <td className="border px-2 py-1">
                                  <table className="w-full border-collapse text-[11px]">
                                    <thead>
                                      <tr className="text-slate-500">
                                        <th className="text-left font-semibold pr-1">Company</th>
                                        <th className="text-right font-semibold">Broiler</th>
                                        <th className="text-right font-semibold">Layer</th>
                                        <th className="text-right font-semibold">Sonali</th>
                                        <th className="text-right font-semibold">Cattle</th>
                                        <th className="text-right font-semibold">Fish</th>
                                        <th className="text-right font-semibold">Total</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {getAgentCompanies(agent).map(
                                        (link: any) => (
                                          <tr key={String(link.company_id)}>
                                            <td className="pr-1 truncate max-w-[90px]">
                                              {link.company_name || "-"}
                                            </td>
                                            <td className="text-right">{Number(link.broiler_sale || 0)}</td>
                                            <td className="text-right">{Number(link.layer_sale || 0)}</td>
                                            <td className="text-right">{Number(link.sonali_sale || 0)}</td>
                                            <td className="text-right">{Number(link.cattle_sale || 0)}</td>
                                            <td className="text-right">{Number(link.fish_sale || 0)}</td>
                                            <td className="text-right font-bold text-emerald-700">
                                              {Number(
                                                (link.monthly_sale ?? linkSaleTotal(link)) || 0
                                              ).toLocaleString()}
                                            </td>
                                          </tr>
                                        )
                                      )}
                                      <tr className="font-bold border-t">
                                        <td className="pr-1">Total</td>
                                        <td colSpan={5}></td>
                                        <td className="text-right text-emerald-700">
                                          {getAgentTotalSale(agent).toLocaleString()}
                                        </td>
                                      </tr>
                                    </tbody>
                                  </table>
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </Popup>
                    </Marker>
                  );
                })}
            </MapContainer>

            {/* Floating Locate Me + Close Button (Middle Side-by-Side) */}
            <div className="absolute top-3 left-1/2 -translate-x-1/2 z-[500] flex flex-col items-center gap-1.5">
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={getMyLocation}
                  className="flex items-center gap-1 bg-blue-600 hover:bg-blue-700 text-white rounded-full shadow-lg px-2.5 py-1.5 text-xs font-semibold whitespace-nowrap transition"
                >
                  <MapPin size={15} /> Locate Me
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setShowMap(false);
                    setSelectedMapAgentId(null);
                    setSelectedLocation(null);
                  }}
                  title="Close Map"
                  className="flex items-center gap-1 bg-purple-500 hover:bg-red-700 text-white rounded-full shadow-lg px-3 py-1.5 text-xs font-bold whitespace-nowrap transition"
                >
                  <X size={14} /> Close
                </button>
              </div>

              {activeDistanceText && (
                <div className="rounded-full bg-slate-900/90 text-white px-3 py-1 text-[11px] font-bold shadow-md whitespace-nowrap">
                  📍 Distance: {activeDistanceText}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          4. SIMPLE BORDERED TABLE GROUPED BY COMPANY
      ================================================== */}
      {companyGroupedAgents.length === 0 ? (
        <div className="bg-white border border-slate-300 rounded-lg p-8 text-center text-xs text-slate-500">
          No agents found for the selected filter.
        </div>
      ) : (
        <div className="space-y-4">
          {companyGroupedAgents.map((group: any) => (
            <div
              key={group.companyId}
              className="bg-white border border-slate-300 rounded-lg overflow-hidden shadow-sm"
            >
              <div className="bg-slate-100 border-b border-slate-300 px-3 py-2 flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-2">
                  {group.logoUrl && (
                    <img
                      src={group.logoUrl}
                      alt={group.companyName}
                      className="w-6 h-6 object-contain rounded bg-white border border-slate-200 p-0.5"
                    />
                  )}
                  <span className="text-xs sm:text-sm font-bold text-slate-800">
                    Company: {group.companyName}
                  </span>
                </div>

                <div className="flex items-center gap-2 text-xs">
                  <span className="font-semibold bg-white border border-slate-300 px-2 py-0.5 rounded text-slate-700">
                    Agents: {group.agents.length}
                  </span>
                  <span className="font-semibold bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded text-emerald-800">
                    Total Sale: {group.totalSale.toLocaleString()}
                  </span>
                </div>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-xs">
                  <thead className="bg-slate-50">
                    <tr>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        SL
                      </th>
                      {group.isNourish && (
                        <th className="border border-slate-300 px-2 py-1.5 text-left">
                          Agent ID
                        </th>
                      )}
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Agent Name
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Proprietor
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Mobile
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Bazar
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Union
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Upazila
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Sale Details
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-right">
                        Monthly Sale
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Status
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-center">
                        Actions
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {group.agents.map((entry: any, idx: number) => {
                      const agent = entry.agent;
                      const link = entry.link;
                      const breakdown = getSaleBreakdown(link);
                      const agentNonEx = isNonExclusive(agent);
                      return (
                        <tr key={agent.id} className="hover:bg-slate-50">
                          <td className="border border-slate-300 px-2 py-1">
                            {idx + 1}
                          </td>
                          {group.isNourish && (
                            <td className="border border-slate-300 px-2 py-1 font-bold text-emerald-700">
                              {agent.agent_code || "-"}
                            </td>
                          )}
                          <td className="border border-slate-300 px-2 py-1 font-semibold text-slate-800">
                            <span className="truncate">{agent.agent_name}</span>
                            {agentNonEx && (
                              <>
                                <button
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    if (nonExPopupAgentId === agent.id) {
                                      setNonExPopupAgentId(null);
                                      setNonExPopupPos(null);
                                      return;
                                    }
                                    // Fixed position ABOVE the badge, clamped
                                    // to the viewport — never clipped by the
                                    // table's overflow container. Falls
                                    // below the badge when there is no room.
                                    const rect =
                                      e.currentTarget.getBoundingClientRect();
                                    const links = getAgentCompanies(agent);
                                    const estHeight = 48 + links.length * 26;
                                    const width = 224;
                                    const left = Math.max(
                                      8,
                                      Math.min(
                                        rect.left,
                                        window.innerWidth - width - 8
                                      )
                                    );
                                    const top =
                                      rect.top - 8 - estHeight < 8
                                        ? rect.bottom + 8
                                        : Math.max(8, rect.top - 8 - estHeight);
                                    setNonExPopupAgentId(agent.id);
                                    setNonExPopupPos({ top, left });
                                  }}
                                  title="Non-exclusive — also works with other companies"
                                  className="ml-1 align-middle rounded bg-violet-100 px-1 py-0 text-[10px] font-bold italic text-violet-700 ring-1 ring-violet-300 hover:bg-violet-200 cursor-pointer"
                                >
                                  non-ex.
                                </button>
                                                              </>
                            )}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {agent.proprietor || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {agent.mobile || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {agent.bazar || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {agent.union_name || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {agent.upazila || "-"}
                          </td>
                          <td
                            className="border border-slate-300 px-2 py-1 whitespace-nowrap"
                            title={`Broiler ${Number(link.broiler_sale || 0)} • Layer ${Number(link.layer_sale || 0)} • Sonali ${Number(link.sonali_sale || 0)} • Cattle ${Number(link.cattle_sale || 0)} • Fish ${Number(link.fish_sale || 0)}`}
                          >
                            {breakdown.length === 0 ? (
                              "-"
                            ) : (
                              <span className="font-medium text-slate-700">
                                <span title="Broiler">B:{Number(link.broiler_sale || 0)}</span>{" "}
                                <span title="Layer">L:{Number(link.layer_sale || 0)}</span>{" "}
                                <span title="Sonali">S:{Number(link.sonali_sale || 0)}</span>{" "}
                                <span title="Cattle">C:{Number(link.cattle_sale || 0)}</span>{" "}
                                <span title="Fish">F:{Number(link.fish_sale || 0)}</span>
                              </span>
                            )}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 text-right font-bold text-slate-800">
                            {Number(
                              (link.monthly_sale ?? linkSaleTotal(link)) || 0
                            ).toLocaleString()}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {agent.status}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 text-center whitespace-nowrap space-x-1">
                            <button
                              onClick={() => setSelectedAgent(agent)}
                              className="bg-blue-600 hover:bg-blue-700 text-white px-2 py-0.5 rounded text-[11px] cursor-pointer"
                            >
                              View
                            </button>
                            <button
                              onClick={() => handleViewAgentOnMap(agent)}
                              className="inline-flex items-center gap-1 bg-teal-600 hover:bg-teal-700 text-white px-2 py-0.5 rounded text-[11px] cursor-pointer"
                            >
                              <MapPin size={10} className="shrink-0" />
                              <span>Map</span>
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

      {/* Non-exclusive companies popover — fixed positioned ABOVE the
          clicked badge so the table's overflow container can never clip
          it. Rendered at the root (outside the scrolling table div),
          with a backdrop click + Escape to close. */}
      {nonExPopupAgent !== null && nonExPopupPos !== null && (
        <>
          <button
            type="button"
            aria-label="Close"
            onClick={() => {
              setNonExPopupAgentId(null);
              setNonExPopupPos(null);
            }}
            className="fixed inset-0 z-[2999] cursor-default bg-transparent"
          />
          <div
            className="fixed z-[3000] w-56 rounded-lg border border-slate-200 bg-white p-2 shadow-xl text-left"
            style={{ top: nonExPopupPos.top, left: nonExPopupPos.left }}
          >
            <div className="flex items-center justify-between mb-1">
              <span className="text-[11px] font-bold text-slate-700">
                Also with
              </span>
              <button
                type="button"
                onClick={() => {
                  setNonExPopupAgentId(null);
                  setNonExPopupPos(null);
                }}
                className="text-slate-400 hover:text-slate-700"
                title="Close"
              >
                <X size={12} />
              </button>
            </div>
            {getAgentCompanies(nonExPopupAgent).map((c: any) => (
              <div
                key={String(c.company_id)}
                className="flex items-center gap-1.5 py-0.5"
              >
                {c.logo_url ? (
                  <img
                    src={c.logo_url}
                    alt=""
                    className="w-4 h-4 object-contain rounded bg-white border border-slate-200 shrink-0"
                  />
                ) : null}
                <span className="truncate text-[11px] text-slate-700">
                  {c.company_name || "-"}
                </span>
                <span className="ml-auto text-[11px] font-bold text-emerald-700 whitespace-nowrap">
                  {Number(
                    (c.monthly_sale ?? linkSaleTotal(c)) || 0
                  ).toLocaleString()}
                </span>
              </div>
            ))}
          </div>
        </>
      )}

      {/* ==================================================
          5. VIEW AGENT MODAL
      ================================================== */}
      {selectedAgent && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[2000] p-2 sm:p-4">
          <div className="bg-white rounded-xl w-full max-w-xl max-h-[92vh] flex flex-col overflow-hidden shadow-2xl border border-slate-300">
            <div className="flex items-center justify-between px-3 py-2 border-b border-slate-300 bg-slate-100">
              <h2 className="text-sm sm:text-base font-bold text-slate-800">
                Agent: {selectedAgent.agent_name}
              </h2>
              <button
                onClick={() => setSelectedAgent(null)}
                className="text-slate-500 hover:text-slate-800"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-3 sm:p-4 space-y-2.5 overflow-y-auto">
              <div className="bg-emerald-50 border border-emerald-300 rounded-lg p-3 space-y-2">
                <div className="flex items-center justify-between text-xs">
                  <span className="font-bold text-emerald-900">
                    📝 Agent Tracking
                  </span>
                  <span className="text-emerald-700 flex items-center gap-1.5">
                    <span>User:</span>
                    <UserAvatar
                      name={currentUserName}
                      avatarUrl={getUserAvatar(undefined, currentUserName)}
                      size="w-4 h-4"
                      textSize="text-[8px]"
                    />
                    <b>{currentUserName}</b>
                  </span>
                </div>

                <div className="max-h-36 overflow-y-auto space-y-1.5">
                  {loadingNotes ? (
                    <div className="text-xs text-slate-500">Loading...</div>
                  ) : agentNotes.length === 0 ? (
                    <div className="text-xs text-slate-400 bg-white border border-emerald-200 rounded p-2 text-center">
                      No tracking notes yet.
                    </div>
                  ) : (
                    agentNotes.map((item) => (
                      <div
                        key={item.id}
                        className="bg-white border border-emerald-200 rounded px-2.5 py-1.5 text-xs flex justify-between items-start gap-2"
                      >
                        <div className="flex items-start gap-1.5 min-w-0 flex-1">
                          <div className="mt-0.5 shrink-0">
                            <UserAvatar
                              name={item.user_name || "Officer"}
                              avatarUrl={getUserAvatar(
                                item.user_id,
                                item.user_name
                              )}
                              size="w-4 h-4"
                              textSize="text-[8px]"
                            />
                          </div>
                          <div className="min-w-0 flex-1 break-words">
                            <span className="font-bold text-emerald-800">
                              {item.user_name}:
                            </span>{" "}
                            <span className="text-slate-800">
                              {renderNoteWithMentions(item.note)}
                            </span>
                          </div>
                        </div>
                        <span className="text-[10px] text-slate-400 shrink-0">
                          {new Date(item.created_at).toLocaleDateString(
                            "en-GB"
                          )}
                        </span>
                      </div>
                    ))
                  )}
                </div>

                <form
                  onSubmit={(e) => handleAddAgentNote(e, selectedAgent.id)}
                  className="flex gap-1.5"
                >
                  <input
                    type="text"
                    placeholder="Write tracking note..."
                    value={newNoteText}
                    onChange={(e) => setNewNoteText(e.target.value)}
                    className="flex-1 rounded border border-emerald-300 bg-white px-2.5 py-1.5 text-xs outline-none"
                  />
                  <button
                    type="submit"
                    disabled={savingNote || !newNoteText.trim()}
                    className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white px-3 py-1.5 rounded text-xs font-semibold flex items-center gap-1 cursor-pointer"
                  >
                    <Send size={12} /> {savingNote ? "Saving..." : "Track"}
                  </button>
                </form>
              </div>

              <table className="w-full border-collapse border border-slate-300 text-xs">
                <tbody>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold w-1/3">
                      Company
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {getAgentCompanies(selectedAgent)
                        .map((c: any) => c.company_name)
                        .filter(Boolean)
                        .join(", ") || "-"}
                      {isNonExclusive(selectedAgent) && (
                        <span className="ml-2 rounded bg-violet-100 px-1.5 py-0.5 text-[10px] font-bold italic text-violet-700">
                          non-exclusive
                        </span>
                      )}
                    </td>
                  </tr>
                  {isNourishName(selectedAgent.companies?.company_name) && (
                    <tr>
                      <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                        Agent ID (Nourish)
                      </td>
                      <td className="border border-slate-300 px-3 py-1.5 font-bold text-emerald-700">
                        {selectedAgent.agent_code || "-"}
                      </td>
                    </tr>
                  )}
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Agent Name
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-bold">
                      {selectedAgent.agent_name}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Proprietor
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedAgent.proprietor || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Mobile
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedAgent.mobile || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Address
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {[
                        selectedAgent.bazar,
                        selectedAgent.union_name,
                        selectedAgent.upazila,
                        selectedAgent.district,
                      ]
                        .filter(Boolean)
                        .join(", ")}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Sale Breakdown
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-semibold text-blue-700">
                      <span>Broiler: {selectedAgent.broiler_sale || 0}</span> •{" "}
                      <span>Layer: {selectedAgent.layer_sale || 0}</span> •{" "}
                      <span>Sonali: {selectedAgent.sonali_sale || 0}</span> •{" "}
                      <span>Cattle: {selectedAgent.cattle_sale || 0}</span> •{" "}
                      <span>Fish: {selectedAgent.fish_sale || 0}</span>
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Total Monthly Sale
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-bold text-emerald-700">
                      {Number(selectedAgent.monthly_sale || 0).toLocaleString()}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Focus Item
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedAgent.focus_item || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Status
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedAgent.status || "Active"}
                    </td>
                  </tr>
                </tbody>
              </table>

              {/* Companies & sales — one row per linked company */}
              <div className="mt-1">
                <div className="text-xs font-bold text-slate-700 mb-1">
                  Companies & Sales
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse border border-slate-300 text-xs">
                    <thead className="bg-slate-50">
                      <tr>
                        <th className="border px-2 py-1 text-left">Company</th>
                        <th className="border px-2 py-1 text-right">Broiler</th>
                        <th className="border px-2 py-1 text-right">Layer</th>
                        <th className="border px-2 py-1 text-right">Sonali</th>
                        <th className="border px-2 py-1 text-right">Cattle</th>
                        <th className="border px-2 py-1 text-right">Fish</th>
                        <th className="border px-2 py-1 text-right">Total</th>
                      </tr>
                    </thead>
                    <tbody>
                      {getAgentCompanies(selectedAgent).map((c: any) => (
                        <tr key={String(c.company_id)}>
                          <td className="border px-2 py-1 font-medium truncate max-w-[140px]">
                            {c.logo_url ? (
                              <img src={c.logo_url} alt="" className="inline w-4 h-4 object-contain mr-1 align-middle" />
                            ) : null}
                            {c.company_name || "-"}
                          </td>
                          <td className="border px-2 py-1 text-right">{Number(c.broiler_sale || 0)}</td>
                          <td className="border px-2 py-1 text-right">{Number(c.layer_sale || 0)}</td>
                          <td className="border px-2 py-1 text-right">{Number(c.sonali_sale || 0)}</td>
                          <td className="border px-2 py-1 text-right">{Number(c.cattle_sale || 0)}</td>
                          <td className="border px-2 py-1 text-right">{Number(c.fish_sale || 0)}</td>
                          <td className="border px-2 py-1 text-right font-bold text-emerald-700">
                            {Number((c.monthly_sale ?? linkSaleTotal(c)) || 0).toLocaleString()}
                          </td>
                        </tr>
                      ))}
                      <tr className="bg-slate-50 font-bold">
                        <td className="border px-2 py-1">Total</td>
                        <td className="border px-2 py-1 text-right">{getAgentCompanies(selectedAgent).reduce((s: number, c: any) => s + Number(c.broiler_sale || 0), 0)}</td>
                        <td className="border px-2 py-1 text-right">{getAgentCompanies(selectedAgent).reduce((s: number, c: any) => s + Number(c.layer_sale || 0), 0)}</td>
                        <td className="border px-2 py-1 text-right">{getAgentCompanies(selectedAgent).reduce((s: number, c: any) => s + Number(c.sonali_sale || 0), 0)}</td>
                        <td className="border px-2 py-1 text-right">{getAgentCompanies(selectedAgent).reduce((s: number, c: any) => s + Number(c.cattle_sale || 0), 0)}</td>
                        <td className="border px-2 py-1 text-right">{getAgentCompanies(selectedAgent).reduce((s: number, c: any) => s + Number(c.fish_sale || 0), 0)}</td>
                        <td className="border px-2 py-1 text-right text-emerald-700">
                          {getAgentTotalSale(selectedAgent).toLocaleString()}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Edit/Delete live ONLY here in the View modal (FarmersPage pattern). */}
              {canEditAgent(selectedAgent) && (
                <div className="flex justify-end gap-2 pt-1">
                  <button
                    onClick={() => {
                      openEditForm(selectedAgent);
                      setSelectedAgent(null);
                    }}
                    className="bg-yellow-500 hover:bg-yellow-600 text-white px-3 py-1.5 rounded text-xs font-semibold cursor-pointer"
                  >
                    Edit
                  </button>
                  <button
                    onClick={() => handleDelete(selectedAgent.id)}
                    className="bg-red-600 hover:bg-red-700 text-white px-3 py-1.5 rounded text-xs font-semibold cursor-pointer"
                  >
                    Delete Agent
                  </button>
                </div>
              )}

            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          6. ADD / EDIT AGENT MODAL (WITH NOURISH AGENT ID OPTION)
      ================================================== */}
      {showForm && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[2000] p-2 sm:p-4">
          <div className="bg-white rounded-xl p-3 sm:p-4 w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl">
            <div className="flex justify-between items-center pb-2.5 mb-3 border-b">
              <h2 className="text-base font-bold text-slate-800">
                {editingId ? "Edit Agent" : "Add Agent"}
              </h2>
              <button
                onClick={() => {
                  setShowForm(false);
                  setEditingId(null);
                  setFormCompanies([]);
                  setFormType("exclusive");
                }}
              >
                <X size={18} />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              <div className="sm:col-span-2 flex flex-wrap items-center gap-3 bg-slate-50 border border-slate-200 rounded px-2.5 py-1.5">
                <span className="text-xs font-bold text-slate-600">Type:</span>
                <label className="flex items-center gap-1 text-xs font-semibold text-slate-700 cursor-pointer">
                  <input
                    type="radio"
                    name="agent-type"
                    checked={formType === "exclusive"}
                    onChange={() => handleTypeChange("exclusive")}
                  />
                  Exclusive
                </label>
                <label className="flex items-center gap-1 text-xs font-semibold text-slate-700 cursor-pointer">
                  <input
                    type="radio"
                    name="agent-type"
                    checked={formType === "nonexclusive"}
                    onChange={() => handleTypeChange("nonexclusive")}
                  />
                  Non-exclusive
                </label>
                <span className="ml-auto text-[11px] font-bold text-emerald-700">
                  Agent Total:{" "}
                  {formCompanies
                    .reduce((s: number, c: any) => s + linkSaleTotal(c), 0)
                    .toLocaleString()}
                </span>
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Primary Company *
                </label>
                <SearchableSelect
                  value={form.company_id}
                  onChange={handlePrimaryCompanyChange}
                  placeholder="Select Company"
                  options={companies.map((c) => ({
                    value: String(c.id),
                    label: c.company_name,
                  }))}
                />
              </div>

              {/* Agent ID Field — ONLY SHOWN WHEN NOURISH IS SELECTED */}
              {isFormCompanyNourish && (
                <div>
                  <label className="block text-xs font-bold text-emerald-700 mb-1">
                    Agent ID (Nourish) *
                  </label>
                  <input
                    placeholder="e.g. 1059"
                    value={form.agent_code}
                    onChange={(e) =>
                      setForm({ ...form, agent_code: e.target.value })
                    }
                    className="w-full rounded border border-emerald-400 bg-emerald-50/40 px-2.5 py-1.5 text-xs outline-none focus:border-emerald-600"
                  />
                </div>
              )}

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Agent Name *
                </label>
                <input
                  placeholder="Agent Name"
                  value={form.agent_name}
                  onChange={(e) =>
                    setForm({ ...form, agent_name: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Proprietor
                </label>
                <input
                  placeholder="Proprietor"
                  value={form.proprietor}
                  onChange={(e) =>
                    setForm({ ...form, proprietor: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Mobile
                </label>
                <input
                  placeholder="017XXXXXXXX"
                  value={form.mobile}
                  onChange={(e) =>
                    setForm({ ...form, mobile: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  District *
                </label>
                <SearchableSelect
                  searchable={false}
                  value={form.district}
                  onChange={(v) =>
                    setForm({
                      ...form,
                      district: v,
                      upazila: "",
                      union_name: "",
                      bazar: "",
                    })
                  }
                  placeholder="Select District"
                  options={districts.map((d) => ({
                    value: d.name,
                    label: d.name,
                  }))}
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Upazila *
                </label>
                <SearchableSelect
                  value={form.upazila}
                  onChange={(v) =>
                    setForm({
                      ...form,
                      upazila: v,
                      union_name: "",
                      bazar: "",
                    })
                  }
                  placeholder="Select Upazila"
                  options={upazilas.map((u) => ({
                    value: u.name,
                    label: u.name,
                  }))}
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Union
                </label>
                <SearchableSelect
                  value={form.union_name}
                  onChange={(v) =>
                    setForm({
                      ...form,
                      union_name: v,
                      bazar: "",
                    })
                  }
                  placeholder="Select Union"
                  options={unions.map((u: any) => ({
                    value: u.name,
                    label: u.name,
                  }))}
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Bazar
                </label>
                <SearchableSelect
                  value={form.bazar}
                  onChange={(v) => setForm({ ...form, bazar: v })}
                  placeholder="Select Bazar"
                  options={filteredBazars.map((bazar) => ({
                    value: bazar.name,
                    label: bazar.name,
                  }))}
                />
              </div>
            </div>

            {/* Map Location Trigger */}
            <div className="mb-3 flex flex-wrap items-center gap-2 bg-slate-50 border border-slate-200 rounded p-2.5">
              <button
                type="button"
                onClick={() => {
                  setPickerCenterOverride(null);
                  setShowAgentMap(true);
                }}
                className="inline-flex items-center gap-1.5 bg-blue-600 text-white px-3 py-1.5 rounded text-xs font-semibold"
              >
                <MapPin size={13} /> Select Location from Map
              </button>

              {form.latitude && form.longitude && (
                <span className="text-emerald-700 text-xs font-semibold">
                  📍 Selected ({Number(form.latitude).toFixed(4)},{" "}
                  {Number(form.longitude).toFixed(4)})
                </span>
              )}
            </div>

            {/* Companies & per-company sales (B/S/L/C/F, auto total) */}
            <div className="space-y-2 mb-3">
              {formCompanies.map((comp: any, idx: number) => {
                const compObj = companies.find(
                  (c) => String(c.id) === String(comp.company_id)
                );
                return (
                  <div
                    key={idx}
                    className="bg-slate-50 border border-slate-200 rounded p-2.5"
                  >
                    <div className="flex flex-wrap items-center gap-2 mb-2">
                      {comp.logo_url || compObj?.logo_url ? (
                        <img
                          src={comp.logo_url || compObj?.logo_url}
                          alt=""
                          className="w-5 h-5 object-contain rounded bg-white border border-slate-200"
                        />
                      ) : null}
                      <SearchableSelect
                        className="flex-1 min-w-[140px]"
                        value={comp.company_id}
                        onChange={(v) => {
                          if (
                            formCompanies.some(
                              (c: any, i: number) =>
                                i !== idx &&
                                String(c.company_id) === String(v)
                            )
                          ) {
                            alert("Company already added.");
                            return;
                          }
                          const picked = companies.find(
                            (c) => String(c.id) === String(v)
                          );
                          updateFormCompany(idx, {
                            company_id: v,
                            company_name: picked?.company_name || "",
                            logo_url: picked?.logo_url || "",
                          });
                        }}
                        placeholder="Select Company"
                        options={companies.map((c) => ({
                          value: String(c.id),
                          label: c.company_name,
                        }))}
                      />
                      <span className="text-[11px] font-bold text-emerald-800 whitespace-nowrap">
                        Total: {linkSaleTotal(comp).toLocaleString()}
                      </span>
                      {formCompanies.length > 1 && (
                        <button
                          type="button"
                          onClick={() => removeFormCompany(idx)}
                          className="text-red-600 hover:text-red-800 text-[11px] font-semibold"
                          title="Remove this company"
                        >
                          Remove
                        </button>
                      )}
                    </div>
                    <div className="grid grid-cols-5 gap-1.5">
                      {(
                        [
                          "broiler_sale",
                          "layer_sale",
                          "sonali_sale",
                          "cattle_sale",
                          "fish_sale",
                        ] as const
                      ).map((k) => (
                        <div key={k}>
                          <label className="block text-[10px] font-semibold text-slate-600 mb-0.5 text-center truncate">
                            {SALE_LABELS[k]}
                          </label>
                          <input
                            type="number"
                            min="0"
                            value={(comp as any)[k] ?? ""}
                            onChange={(e) =>
                              updateFormCompany(idx, { [k]: e.target.value })
                            }
                            className="w-full rounded border border-slate-300 bg-white px-1 py-1 text-xs text-center"
                          />
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })}
              {formType === "nonexclusive" && (
                <div className="flex flex-wrap items-center gap-2">
                  <SearchableSelect
                    className="flex-1 min-w-[160px]"
                    value={companyPickerId}
                    onChange={setCompanyPickerId}
                    placeholder="Select company to add…"
                    options={companies
                      .filter(
                        (c) =>
                          !formCompanies.some(
                            (fc: any) => String(fc.company_id) === String(c.id)
                          )
                      )
                      .map((c) => ({
                        value: String(c.id),
                        label: c.company_name,
                      }))}
                  />
                  <button
                    type="button"
                    onClick={() => {
                      addFormCompany(companyPickerId);
                      setCompanyPickerId("");
                    }}
                    disabled={!companyPickerId}
                    className="inline-flex items-center gap-1 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white px-3 py-1.5 rounded text-xs font-semibold"
                  >
                    <Plus size={13} /> Add Company
                  </button>
                </div>
              )}
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Focus Item
                </label>
                <SearchableSelect
                  searchable={false}
                  value={form.focus_item}
                  onChange={(v) => setForm({ ...form, focus_item: v })}
                  placeholder="Select Focus Item"
                  options={[
                    { value: "Broiler Feed", label: "Broiler Feed" },
                    { value: "Layer Feed", label: "Layer Feed" },
                    { value: "Sonali Feed", label: "Sonali Feed" },
                    { value: "Cattle Feed", label: "Cattle Feed" },
                    { value: "Fish Feed", label: "Fish Feed" },
                    { value: "Goat Feed", label: "Goat Feed" },
                    { value: "Duck Feed", label: "Duck Feed" },
                    { value: "Others", label: "Others" },
                  ]}
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Status
                </label>
                <SearchableSelect
                  searchable={false}
                  value={form.status}
                  onChange={(v) => setForm({ ...form, status: v })}
                  placeholder="Select Status"
                  options={[
                    { value: "Active", label: "Active" },
                    { value: "Inactive", label: "Inactive" },
                  ]}
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t">
              <button
                onClick={() => {
                  setShowForm(false);
                  setEditingId(null);
                  setFormCompanies([]);
                  setFormType("exclusive");
                }}
                className="px-4 py-1.5 rounded border border-slate-300 text-xs font-semibold"
              >
                Cancel
              </button>
              <button
                onClick={handleSave}
                disabled={savingAgent}
                className="px-4 py-1.5 rounded bg-green-600 text-white text-xs font-semibold cursor-pointer disabled:opacity-60"
              >
                {savingAgent
                  ? "Saving..."
                  : editingId
                  ? "Update Agent"
                  : "Save Agent"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          7. LOCATION PICKER MODAL WITH LOCATE ME BUTTON
      ================================================== */}
      {showAgentMap && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[2100] p-2 sm:p-4">
          <div className="bg-white p-3 sm:p-4 rounded-xl w-full max-w-4xl shadow-2xl space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="font-bold text-sm sm:text-base">
                Select Agent Location
              </h2>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleLocateMeInPicker}
                  disabled={locatingInPicker}
                  className="inline-flex items-center gap-1 bg-blue-600 hover:bg-blue-700 text-white px-3 py-1.5 rounded text-xs font-semibold"
                >
                  <Navigation size={13} />
                  {locatingInPicker ? "Locating..." : "Locate Me"}
                </button>

                {form.latitude && form.longitude && (
                  <button
                    type="button"
                    onClick={() => setShowAgentMap(false)}
                    className="inline-flex items-center gap-1 bg-green-600 hover:bg-green-700 text-white px-3 py-1.5 rounded text-xs font-semibold"
                  >
                    <CheckCircle2 size={13} /> Use Selected Location
                  </button>
                )}

                <button onClick={() => setShowAgentMap(false)}>
                  <X size={18} />
                </button>
              </div>
            </div>

            <LocationPicker
              key={
                pickerCenterOverride
                  ? `${pickerCenterOverride[0]}_${pickerCenterOverride[1]}`
                  : "default_agent_picker"
              }
              center={
                pickerCenterOverride
                  ? pickerCenterOverride
                  : selectedUpazila
                  ? [Number(selectedUpazila.lat), Number(selectedUpazila.lon)]
                  : selectedDistrict
                  ? [
                      Number(selectedDistrict.lat),
                      Number(selectedDistrict.lon),
                    ]
                  : [24.7471, 90.4203]
              }
              zoom={
                pickerCenterOverride
                  ? 17
                  : selectedUpazila
                  ? 13
                  : selectedDistrict
                  ? 11
                  : 7
              }
              selectedPosition={
                form.latitude && form.longitude
                  ? [Number(form.latitude), Number(form.longitude)]
                  : null
              }
              onSelect={(lat, lng) => {
                setForm({
                  ...form,
                  latitude: lat.toString(),
                  longitude: lng.toString(),
                });
                setShowAgentMap(false);
              }}
              markers={
                <>
                  {agents
                    .filter((a) => a.latitude && a.longitude)
                    .map((agent) => (
                      <Marker
                        key={agent.id}
                        position={[
                          Number(agent.latitude),
                          Number(agent.longitude),
                        ]}
                        icon={getAgentMarkerIcon(agent)}
                      />
                    ))}
                </>
              }
            />
          </div>
        </div>
      )}

      {/* ==================================================
          8. AGENT SUMMARY MODAL
      ================================================== */}
      {showSummary && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[2000] p-2 sm:p-4">
          <div className="bg-white rounded-xl p-3 sm:p-4 w-full max-w-3xl max-h-[92vh] overflow-auto">
            <div className="flex justify-between items-center pb-2 mb-3 border-b">
              <div>
                <h2 className="text-base sm:text-lg font-bold">
                  Agent Summary — {selectedCompanyName}
                </h2>
                <p className="text-xs text-slate-500">
                  District: {districtFilter || "All"} | Upazila:{" "}
                  {upazilaFilter || "All"} | Union: {unionFilter || "All"}
                </p>
              </div>
              <button onClick={() => setShowSummary(false)}>
                <X size={18} />
              </button>
            </div>

            <div className="mb-3 flex items-center gap-4 text-xs sm:text-sm font-semibold">
              <span>Total Agents: {filteredAgents.length}</span>
              <span className="text-violet-700">
                Non-exclusive: {nonExclusiveCount}
              </span>
              <span className="text-emerald-700">
                Total Monthly Sale: {totalMonthlySale.toLocaleString()}
              </span>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-xs">
                <thead className="bg-slate-100">
                  <tr>
                    <th className="border px-2.5 py-1.5 text-left">Company</th>
                    <th className="border px-2 py-1.5 text-center">Agents</th>
                    <th className="border px-2 py-1 text-right">Broiler</th>
                    <th className="border px-2 py-1 text-right">Layer</th>
                    <th className="border px-2 py-1 text-right">Sonali</th>
                    <th className="border px-2 py-1 text-right">Cattle</th>
                    <th className="border px-2 py-1 text-right">Fish</th>
                    <th className="border px-2.5 py-1.5 text-right bg-emerald-50">
                      Total Sale
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {companySummary.map((c: any) => (
                    <tr key={c.name} className="hover:bg-slate-50">
                      <td className="border px-2.5 py-1.5 font-semibold">
                        {c.name}
                      </td>
                      <td className="border px-2 py-1.5 text-center font-bold">
                        {c.agents}
                      </td>
                      <td className="border px-2 py-1.5 text-right">
                        {c.broiler.toLocaleString()}
                      </td>
                      <td className="border px-2 py-1.5 text-right">
                        {c.layer.toLocaleString()}
                      </td>
                      <td className="border px-2 py-1.5 text-right">
                        {c.sonali.toLocaleString()}
                      </td>
                      <td className="border px-2 py-1.5 text-right">
                        {c.cattle.toLocaleString()}
                      </td>
                      <td className="border px-2 py-1.5 text-right">
                        {c.fish.toLocaleString()}
                      </td>
                      <td className="border px-2.5 py-1.5 text-right font-bold text-emerald-700 bg-emerald-50/50">
                        {c.sale.toLocaleString()}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}