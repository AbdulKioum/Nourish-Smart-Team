import React, { useEffect, useState, useRef } from "react";
import LocationPicker from "../../components/maps/LocationPicker";
import { supabase } from "../../lib/supabase";
import {
  getAgents,
  createAgent,
  updateAgent,
  deleteAgent,
} from "../../services/agentService";
import { useAuth } from "../../context/AuthContext";
import {
  MapContainer,
  Marker,
  Popup,
  Tooltip,
  useMap,
} from "react-leaflet";
import L from "leaflet";
import "leaflet/dist/leaflet.css";
import { saveAs } from "file-saver";
import { getBazars } from "../../services/bazarService";
import { createAgentMarker } from "../../utils/mapMarkers";
import MapLayers from "../../components/maps/MapLayers";
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

function isNourishName(name?: string) {
  return Boolean(name && name.toLowerCase().includes("nourish"));
}

function RecenterMap({ position }: { position: [number, number] }) {
  const map = useMap();
  useEffect(() => {
    if (!position) return;
    map.flyTo(position, 17, { animate: true, duration: 1.3 });
  }, [position, map]);
  return null;
}

function FitBounds({ positions }: { positions: [number, number][] }) {
  const map = useMap();
  useEffect(() => {
    if (positions.length > 0) {
      map.fitBounds(positions, { padding: [45, 45] });
    }
  }, [positions, map]);
  return null;
}

export default function AgentsPage() {
  const { role } = useAuth();
  const mapSectionRef = useRef<HTMLDivElement | null>(null);
  const agentMarkerRefs = useRef<Record<number, L.Marker | null>>({});

  if (role === "viewer") {
    return <div className="p-6 text-center text-slate-500">Access Denied</div>;
  }

  const isFullAccess = role === "super_admin" || role === "manager";

  // User Profile & Assigned Areas
  const [currentUserId, setCurrentUserId] = useState("");
  const [currentUserName, setCurrentUserName] = useState("User");
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>([]);

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

  // Filters
  const [search, setSearch] = useState("");
  const [companyFilter, setCompanyFilter] = useState("");
  const [districtFilter, setDistrictFilter] = useState("");
  const [upazilaFilter, setUpazilaFilter] = useState("");
  const [unionFilter, setUnionFilter] = useState("");
  const [bazarFilter, setBazarFilter] = useState("");

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
      const { data: profile } = await supabase
        .from("users")
        .select("full_name, email, role, districts, upazilas")
        .eq("id", user.id)
        .single();

      if (profile) {
        uName = profile.full_name || profile.email || uName;
        uDistricts = parseList(profile.districts);
        uUpazilas = parseList(profile.upazilas);
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
    const allowedAgents = isFullAccess
      ? allAgents
      : allAgents.filter((a: any) => {
          const matchUpa =
            uUpazilas.length === 0 || uUpazilas.includes(a.upazila);
          const matchDist =
            uDistricts.length === 0 || uDistricts.includes(a.district);
          return matchUpa && matchDist;
        });

    setAgents(allowedAgents);
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
    loadAgentNotes(agent.id);

    setTimeout(() => {
      mapSectionRef.current?.scrollIntoView({
        behavior: "smooth",
        block: "center",
      });
      const marker = agentMarkerRefs.current[agent.id];
      if (marker) {
        marker.openPopup();
      }
    }, 350);
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

  // ==================================================
  // SAVE, DELETE & LOCATE ME (INCLUDING NOURISH AGENT ID)
  // ==================================================
  async function handleSave() {
    try {
      if (
        !form.company_id ||
        !form.agent_name.trim() ||
        !form.district ||
        !form.upazila
      ) {
        alert("Company, Agent Name, District and Upazila are required.");
        return;
      }

      const cleanAgentCode = isFormCompanyNourish
        ? form.agent_code.trim()
        : "";

      const payload: any = {
        ...form,
        agent_code: cleanAgentCode,
        company_id: Number(form.company_id),
        broiler_sale: Number(form.broiler_sale || 0),
        layer_sale: Number(form.layer_sale || 0),
        sonali_sale: Number(form.sonali_sale || 0),
        cattle_sale: Number(form.cattle_sale || 0),
        fish_sale: Number(form.fish_sale || 0),
        monthly_sale: form.monthly_sale === "" ? 0 : Number(form.monthly_sale),
        latitude: form.latitude === "" ? null : Number(form.latitude),
        longitude: form.longitude === "" ? null : Number(form.longitude),
      };

      if (editingId) {
        await updateAgent(editingId, payload);
        await supabase
          .from("agents")
          .update({ agent_code: cleanAgentCode })
          .eq("id", editingId);
        alert("Agent Updated");
      } else {
        const created: any = await createAgent(payload);
        const createdId = created?.id || created?.[0]?.id;
        if (createdId) {
          await supabase
            .from("agents")
            .update({ agent_code: cleanAgentCode })
            .eq("id", createdId);
        } else if (cleanAgentCode) {
          await supabase
            .from("agents")
            .update({ agent_code: cleanAgentCode })
            .eq("agent_name", form.agent_name.trim())
            .eq("company_id", Number(form.company_id));
        }
        alert("Agent Saved");
      }

      setShowForm(false);
      setEditingId(null);
      setForm({ ...EMPTY_FORM });
      loadData();
    } catch (error: any) {
      alert(error?.message || "Save Failed");
    }
  }

  async function handleDelete(id: number) {
    if (!confirm("Delete Agent?")) return;
    await deleteAgent(id);
    loadData();
  }

  function handleLocateMeInPicker() {
    if (!navigator.geolocation) {
      alert("Geolocation is not supported by your browser.");
      return;
    }
    setLocatingInPicker(true);
    navigator.geolocation.getCurrentPosition(
      (pos) => {
        const lat = pos.coords.latitude;
        const lng = pos.coords.longitude;
        setPickerCenterOverride([lat, lng]);
        setForm((prev) => ({
          ...prev,
          latitude: lat.toFixed(6),
          longitude: lng.toFixed(6),
        }));
        setLocatingInPicker(false);
      },
      () => {
        setLocatingInPicker(false);
        alert("Unable to retrieve your location. Please allow GPS access.");
      },
      { enableHighAccuracy: true }
    );
  }

  function getDistanceKm(
    lat1: number,
    lon1: number,
    lat2: number,
    lon2: number
  ) {
    const R = 6371;
    const dLat = ((lat2 - lat1) * Math.PI) / 180;
    const dLon = ((lon2 - lon1) * Math.PI) / 180;
    const a =
      Math.sin(dLat / 2) * Math.sin(dLat / 2) +
      Math.cos((lat1 * Math.PI) / 180) *
        Math.cos((lat2 * Math.PI) / 180) *
        Math.sin(dLon / 2) *
        Math.sin(dLon / 2);
    const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
    return Math.round(R * c * 100) / 100;
  }

  function getMyLocation() {
    if (!navigator.geolocation) {
      alert("Geolocation not supported");
      return;
    }
    navigator.geolocation.getCurrentPosition(
      (position) => {
        const coords: [number, number] = [
          position.coords.latitude,
          position.coords.longitude,
        ];
        setUserLocation(coords);
        setSelectedLocation(coords);
        setMyLocationMode(true);

        const nearest = filteredAgents
          .filter((a) => a.latitude && a.longitude)
          .map((a) => ({
            ...a,
            distance: getDistanceKm(
              coords[0],
              coords[1],
              Number(a.latitude),
              Number(a.longitude)
            ),
          }))
          .sort((a, b) => a.distance - b.distance)
          .slice(0, 10);

        setNearestAgents(nearest);
      },
      () => {
        alert("Location access denied");
      }
    );
  }

  // ==================================================
  // FILTERED AGENTS & COMPANY-WISE GROUPING
  // ==================================================
  const filteredAgents = agents.filter((agent) => {
    const q = search.toLowerCase();
    const matchSearch =
      !q ||
      agent.agent_name?.toLowerCase().includes(q) ||
      agent.agent_code?.toLowerCase().includes(q) ||
      agent.proprietor?.toLowerCase().includes(q) ||
      agent.mobile?.toLowerCase().includes(q) ||
      agent.bazar?.toLowerCase().includes(q);

    const matchCompany =
      companyFilter === "" || String(agent.company_id) === companyFilter;
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
      matchDistrict &&
      matchUpazila &&
      matchUnion &&
      matchBazar
    );
  });

  const sortedAgents = [...filteredAgents].sort((a: any, b: any) => {
    const companyCompare = (
      a.companies?.company_name || ""
    ).localeCompare(b.companies?.company_name || "");
    if (companyCompare !== 0) return companyCompare;
    return (a.agent_name || "").localeCompare(b.agent_name || "");
  });

  const companyGroupedAgents = Object.values(
    sortedAgents.reduce((acc: Record<string, any>, agent: any) => {
      const key = String(agent.company_id || "unassigned");
      if (!acc[key]) {
        const compName =
          agent.companies?.company_name || "Unassigned Company";
        acc[key] = {
          companyId: key,
          companyName: compName,
          isNourish: isNourishName(compName),
          logoUrl: agent.companies?.logo_url || "",
          totalSale: 0,
          agents: [],
        };
      }
      acc[key].agents.push(agent);
      acc[key].totalSale += Number(agent.monthly_sale || 0);
      return acc;
    }, {})
  );

  const selectedCompanyName =
    companies.find((c) => String(c.id) === companyFilter)?.company_name ||
    "All Companies";

  const totalMonthlySale = filteredAgents.reduce(
    (sum, agent) => sum + Number(agent.monthly_sale || 0),
    0
  );

  const companySummary = Object.values(
    filteredAgents.reduce((acc: any, agent) => {
      const company = agent.companies?.company_name || "Unknown";
      if (!acc[company]) {
        acc[company] = {
          name: company,
          agents: 0,
          broiler: 0,
          layer: 0,
          sonali: 0,
          cattle: 0,
          fish: 0,
          sale: 0,
        };
      }
      acc[company].agents += 1;
      acc[company].broiler += Number(agent.broiler_sale || 0);
      acc[company].layer += Number(agent.layer_sale || 0);
      acc[company].sonali += Number(agent.sonali_sale || 0);
      acc[company].cattle += Number(agent.cattle_sale || 0);
      acc[company].fish += Number(agent.fish_sale || 0);
      acc[company].sale += Number(agent.monthly_sale || 0);
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
        ]);
      });

      for (let col = 1; col <= 17; col++) {
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
      ];
      widths.forEach((w, idx) => (worksheet.getColumn(idx + 1).width = w));
      worksheet.getRow(1).height = 30;
      worksheet.views = [{ state: "frozen", ySplit: 1 }];
      worksheet.autoFilter = { from: "A1", to: "Q1" };

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const today = new Date().toISOString().split("T")[0];
      saveAs(
        blob,
        `Agents_${selectedCompanyName.replace(/\s+/g, "_")}_${today}.xlsx`
      );
    } catch (error) {
      console.error("Agent Excel Export Error:", error);
      alert("Agent Excel export failed.");
    }
  }

  function openEditForm(agent: any) {
    setEditingId(agent.id);
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
      <div className="bg-white border border-slate-300 rounded-xl p-2.5 shadow-sm">
        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2 items-center">
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
              className="w-full rounded-lg border border-slate-300 pl-8 pr-2.5 py-1.5 text-xs outline-none focus:border-blue-500"
            />
          </div>

          <select
            value={districtFilter}
            onChange={(e) => {
              setDistrictFilter(e.target.value);
              setUpazilaFilter("");
              setUnionFilter("");
              setBazarFilter("");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Districts</option>
            {districts.map((d) => (
              <option key={d.id} value={d.name}>
                {d.name}
              </option>
            ))}
          </select>

          <select
            value={upazilaFilter}
            onChange={(e) => {
              setUpazilaFilter(e.target.value);
              setUnionFilter("");
              setBazarFilter("");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Upazilas</option>
            {filterUpazilas.map((u: any) => (
              <option key={u.id} value={u.name}>
                {u.name}
              </option>
            ))}
          </select>

          <select
            value={unionFilter}
            onChange={(e) => {
              setUnionFilter(e.target.value);
              setBazarFilter("");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Unions</option>
            {filterUnions.map((u: any) => (
              <option key={u.id} value={u.name}>
                {u.name}
              </option>
            ))}
          </select>

          <select
            value={bazarFilter}
            onChange={(e) => setBazarFilter(e.target.value)}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Bazars</option>
            {filterBazars.map((b: any) => (
              <option key={b.id} value={b.name}>
                {b.name}
              </option>
            ))}
          </select>

          <div className="col-span-2 sm:col-span-1 flex items-center gap-1.5">
            <select
              value={companyFilter}
              onChange={(e) => setCompanyFilter(e.target.value)}
              className="w-full rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
            >
              <option value="">All Companies</option>
              {companies.map((c) => (
                <option key={c.id} value={c.id}>
                  {c.company_name}
                </option>
              ))}
            </select>

            <button
              onClick={clearFilters}
              title="Reset Filters"
              className="shrink-0 inline-flex items-center justify-center rounded-lg bg-slate-100 hover:bg-slate-200 border border-slate-300 px-2.5 py-1.5 text-xs font-semibold text-slate-700 transition"
            >
              <RotateCcw size={13} />
            </button>
          </div>
        </div>
      </div>

      {/* ==================================================
          2. ADD AGENT, SUMMARY, EXPORT, VIEW MAP BAR
      ================================================== */}
      <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-300 rounded-xl px-3.5 py-2 shadow-sm">
        <div className="flex flex-wrap items-center gap-2">
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
              setShowForm(true);
            }}
            className="inline-flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <Plus size={14} /> Add Agent
          </button>

          <button
            onClick={() => {
              setShowMap(false);
              setShowSummary(true);
            }}
            className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <BarChart3 size={14} /> Summary
          </button>

          <button
            onClick={exportAgentsExcel}
            className="inline-flex items-center gap-1.5 bg-emerald-700 hover:bg-emerald-800 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <FileSpreadsheet size={14} /> Export Excel
          </button>

          <button
            onClick={() => {
              if (!showMap) {
                setMyLocationMode(false);
                setUserLocation(null);
                setSelectedLocation(null);
              }
              setShowSummary(false);
              setShowMap(!showMap);
            }}
            className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition ${
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
          className="bg-white rounded-xl shadow border border-slate-300 p-3"
        >
          <h3 className="text-sm font-bold text-slate-800 mb-2">
            Agent Locations
          </h3>

          <div className="w-full overflow-hidden rounded-xl border border-slate-300 relative">
            <MapContainer
              center={mapCenter}
              zoom={8}
              style={{ height: "480px", width: "100%" }}
            >
              <MapLayers />

              {!myLocationMode && !selectedLocation && (
                <FitBounds
                  positions={filteredAgents
                    .filter((a) => a.latitude && a.longitude)
                    .map((a) => [Number(a.latitude), Number(a.longitude)])}
                />
              )}

              {selectedLocation && <RecenterMap position={selectedLocation} />}

              {userLocation && (
                <Marker position={userLocation}>
                  <Tooltip permanent direction="top" offset={[0, -15]}>
                    📍 My Location
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
                              onClick={() => {
                                setSelectedLocation([
                                  Number(a.latitude),
                                  Number(a.longitude),
                                ]);
                              }}
                              className="text-blue-600 font-bold hover:underline text-left"
                            >
                              {index + 1}. {a.agent_name}
                            </button>
                            <div className="text-slate-500">
                              {a.distance} km away
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </Popup>
                </Marker>
              )}

              {filteredAgents
                .filter((a) => a.latitude && a.longitude)
                .map((agent) => {
                  const breakdown = getSaleBreakdown(agent);

                  return (
                    <Marker
                      key={agent.id}
                      ref={(ref) => {
                        agentMarkerRefs.current[agent.id] = ref;
                      }}
                      position={[
                        Number(agent.latitude),
                        Number(agent.longitude),
                      ]}
                      icon={createAgentMarker(
                        agent.companies?.logo_url || "",
                        agent.agent_name || ""
                      )}
                      eventHandlers={{
                        click: () => {
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
                                {agent.companies?.company_name}
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
                                    className="bg-white border border-emerald-100 rounded px-2 py-1 text-[11px]"
                                  >
                                    <span className="font-bold text-emerald-800">
                                      {n.user_name}:
                                    </span>{" "}
                                    <span className="text-slate-700">
                                      {n.note}
                                    </span>
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
                                className="bg-emerald-600 text-white px-2 py-1 rounded text-[11px] font-semibold"
                              >
                                Save
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
                                  Sale Breakdown
                                </td>
                                <td className="border px-2 py-1 font-semibold text-blue-700">
                                  {breakdown.length === 0
                                    ? "-"
                                    : breakdown
                                        .map((b) => `${b.label}: ${b.val}`)
                                        .join(", ")}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Monthly Sale
                                </td>
                                <td className="border px-2 py-1 font-bold text-emerald-700">
                                  {Number(
                                    agent.monthly_sale || 0
                                  ).toLocaleString()}
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

            <button
              onClick={getMyLocation}
              className="absolute top-3 left-1/2 -translate-x-1/2 z-[1000] flex items-center gap-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-full shadow-lg px-3.5 py-1.5 text-xs font-semibold transition"
            >
              <MapPin size={15} /> Locate Me
            </button>
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
                    {group.agents.map((agent: any, idx: number) => {
                      const breakdown = getSaleBreakdown(agent);
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
                            {agent.agent_name}
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
                          <td className="border border-slate-300 px-2 py-1">
                            {breakdown.length === 0
                              ? "-"
                              : breakdown
                                  .map((b) => `${b.label}: ${b.val}`)
                                  .join(", ")}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 text-right font-bold text-slate-800">
                            {Number(agent.monthly_sale || 0).toLocaleString()}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {agent.status}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 text-center whitespace-nowrap space-x-1">
                            <button
                              onClick={() => setSelectedAgent(agent)}
                              className="bg-blue-600 hover:bg-blue-700 text-white px-2 py-0.5 rounded text-[11px]"
                            >
                              View
                            </button>
                            <button
                              onClick={() => handleViewAgentOnMap(agent)}
                              className="bg-teal-600 hover:bg-teal-700 text-white px-2 py-0.5 rounded text-[11px]"
                            >
                              View on Map
                            </button>
                            <button
                              onClick={() => openEditForm(agent)}
                              className="bg-yellow-500 hover:bg-yellow-600 text-white px-2 py-0.5 rounded text-[11px]"
                            >
                              Edit
                            </button>
                            <button
                              onClick={() => handleDelete(agent.id)}
                              className="bg-red-600 hover:bg-red-700 text-white px-2 py-0.5 rounded text-[11px]"
                            >
                              Delete
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

      {/* ==================================================
          5. VIEW AGENT MODAL
      ================================================== */}
      {selectedAgent && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-4">
          <div className="bg-white rounded-xl w-full max-w-xl max-h-[90vh] flex flex-col overflow-hidden shadow-2xl border border-slate-300">
            <div className="flex items-center justify-between px-4 py-2.5 border-b border-slate-300 bg-slate-100">
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

            <div className="p-4 space-y-3 overflow-y-auto">
              <div className="bg-emerald-50 border border-emerald-300 rounded-lg p-3 space-y-2">
                <div className="flex items-center justify-between text-xs">
                  <span className="font-bold text-emerald-900">
                    📝 Agent Tracking
                  </span>
                  <span className="text-emerald-700">
                    User: <b>{currentUserName}</b>
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
                        className="bg-white border border-emerald-200 rounded px-2.5 py-1.5 text-xs flex justify-between gap-2"
                      >
                        <div>
                          <span className="font-bold text-emerald-800">
                            {item.user_name}:
                          </span>{" "}
                          <span className="text-slate-800">{item.note}</span>
                        </div>
                        <span className="text-[10px] text-slate-400 shrink-0">
                          {new Date(item.created_at).toLocaleDateString("en-GB")}
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
                    className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white px-3 py-1.5 rounded text-xs font-semibold flex items-center gap-1"
                  >
                    <Send size={12} /> Post
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
                      {selectedAgent.companies?.company_name || "-"}
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
                      Broiler: {selectedAgent.broiler_sale || 0} | Layer:{" "}
                      {selectedAgent.layer_sale || 0} | Sonali:{" "}
                      {selectedAgent.sonali_sale || 0} | Cattle:{" "}
                      {selectedAgent.cattle_sale || 0} | Fish:{" "}
                      {selectedAgent.fish_sale || 0}
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
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          6. ADD / EDIT AGENT MODAL (WITH NOURISH AGENT ID OPTION)
      ================================================== */}
      {showForm && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-4">
          <div className="bg-white rounded-xl p-4 sm:p-5 w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl">
            <div className="flex justify-between items-center pb-2.5 mb-3 border-b">
              <h2 className="text-base font-bold text-slate-800">
                {editingId ? "Edit Agent" : "Add Agent"}
              </h2>
              <button
                onClick={() => {
                  setShowForm(false);
                  setEditingId(null);
                }}
              >
                <X size={18} />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Company *
                </label>
                <select
                  value={form.company_id}
                  onChange={(e) =>
                    setForm({ ...form, company_id: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Company</option>
                  {companies.map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.company_name}
                    </option>
                  ))}
                </select>
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
                <select
                  value={form.district}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      district: e.target.value,
                      upazila: "",
                      union_name: "",
                      bazar: "",
                    })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select District</option>
                  {districts.map((d) => (
                    <option key={d.id} value={d.name}>
                      {d.name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Upazila *
                </label>
                <select
                  value={form.upazila}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      upazila: e.target.value,
                      union_name: "",
                      bazar: "",
                    })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Upazila</option>
                  {upazilas.map((u) => (
                    <option key={u.id} value={u.name}>
                      {u.name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Union
                </label>
                <select
                  value={form.union_name}
                  onChange={(e) =>
                    setForm({
                      ...form,
                      union_name: e.target.value,
                      bazar: "",
                    })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Union</option>
                  {unions.map((u: any) => (
                    <option key={u.id} value={u.name}>
                      {u.name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Bazar
                </label>
                <select
                  value={form.bazar}
                  onChange={(e) =>
                    setForm({ ...form, bazar: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Bazar</option>
                  {filteredBazars.map((bazar) => (
                    <option key={bazar.id} value={bazar.name}>
                      {bazar.name}
                    </option>
                  ))}
                </select>
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

            {/* Sales Inputs */}
            <div className="grid grid-cols-2 sm:grid-cols-3 gap-2.5 mb-3 bg-slate-50 border border-slate-200 rounded p-2.5">
              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Broiler Sale
                </label>
                <input
                  type="number"
                  value={form.broiler_sale}
                  onChange={(e) =>
                    setForm({ ...form, broiler_sale: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Layer Sale
                </label>
                <input
                  type="number"
                  value={form.layer_sale}
                  onChange={(e) =>
                    setForm({ ...form, layer_sale: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Sonali Sale
                </label>
                <input
                  type="number"
                  value={form.sonali_sale}
                  onChange={(e) =>
                    setForm({ ...form, sonali_sale: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Cattle Sale
                </label>
                <input
                  type="number"
                  value={form.cattle_sale}
                  onChange={(e) =>
                    setForm({ ...form, cattle_sale: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-[11px] font-semibold text-slate-600 mb-1">
                  Fish Sale
                </label>
                <input
                  type="number"
                  value={form.fish_sale}
                  onChange={(e) =>
                    setForm({ ...form, fish_sale: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-[11px] font-bold text-emerald-800 mb-1">
                  Total Monthly Sale
                </label>
                <input
                  type="number"
                  value={form.monthly_sale}
                  readOnly
                  className="w-full rounded border border-emerald-300 bg-emerald-50 px-2 py-1 text-xs font-bold text-emerald-900"
                />
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Focus Item
                </label>
                <select
                  value={form.focus_item}
                  onChange={(e) =>
                    setForm({ ...form, focus_item: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Focus Item</option>
                  <option value="Broiler Feed">Broiler Feed</option>
                  <option value="Layer Feed">Layer Feed</option>
                  <option value="Sonali Feed">Sonali Feed</option>
                  <option value="Cattle Feed">Cattle Feed</option>
                  <option value="Fish Feed">Fish Feed</option>
                  <option value="Goat Feed">Goat Feed</option>
                  <option value="Duck Feed">Duck Feed</option>
                  <option value="Others">Others</option>
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Status
                </label>
                <select
                  value={form.status}
                  onChange={(e) =>
                    setForm({ ...form, status: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="Active">Active</option>
                  <option value="Inactive">Inactive</option>
                </select>
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t">
              <button
                onClick={() => {
                  setShowForm(false);
                  setEditingId(null);
                }}
                className="px-4 py-1.5 rounded border border-slate-300 text-xs font-semibold"
              >
                Cancel
              </button>
              <button
                onClick={handleSave}
                className="px-4 py-1.5 rounded bg-green-600 text-white text-xs font-semibold"
              >
                {editingId ? "Update Agent" : "Save Agent"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ==================================================
          7. LOCATION PICKER MODAL WITH LOCATE ME BUTTON
      ================================================== */}
      {showAgentMap && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[60] p-2 sm:p-4">
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
                        icon={createAgentMarker(
                          agent.companies?.logo_url,
                          agent.agent_name
                        )}
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
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-2 sm:p-4">
          <div className="bg-white rounded-xl p-4 sm:p-6 w-full max-w-3xl max-h-[85vh] overflow-auto">
            <div className="flex justify-between items-center pb-3 mb-4 border-b">
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
                    <th className="border px-2 py-1.5 text-right">Broiler</th>
                    <th className="border px-2 py-1.5 text-right">Layer</th>
                    <th className="border px-2 py-1.5 text-right">Sonali</th>
                    <th className="border px-2 py-1.5 text-right">Cattle</th>
                    <th className="border px-2 py-1.5 text-right">Fish</th>
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