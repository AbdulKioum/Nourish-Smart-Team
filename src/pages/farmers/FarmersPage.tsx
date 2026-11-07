import React, { useEffect, useState, useRef } from "react";
import { supabase } from "../../lib/supabase";
import LocationPicker from "../../components/maps/LocationPicker";
import {
  getFarmers,
  createFarmer,
  updateFarmer,
  deleteFarmer,
} from "../../services/farmerService";
import { useAuth } from "../../context/AuthContext";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import {
  MapContainer,
  Marker,
  Popup,
  Tooltip,
  useMap,
} from "react-leaflet";
import L from "leaflet";
import "leaflet/dist/leaflet.css";
import MapLayers from "../../components/maps/MapLayers";
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
import ExcelJS from "exceljs";

delete (L.Icon.Default.prototype as any)._getIconUrl;

L.Icon.Default.mergeOptions({
  iconRetinaUrl:
    "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon-2x.png",
  iconUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-icon.png",
  shadowUrl: "https://unpkg.com/leaflet@1.9.4/dist/images/marker-shadow.png",
});

const EMPTY_FORM = {
  company_id: "",
  agent_id: "",
  farmer_name: "",
  mobile: "",
  village: "",
  union_name: "",
  upazila: "",
  district: "",
  farm_category: "",
  broiler: "",
  layer: "",
  sonali: "",
  duck: "",
  quail: "",
  deshi_chicken: "",
  dairy_cow: "",
  bull: "",
  calf_heifer: "",
  pangus: "",
  tilapia: "",
  koi: "",
  shing: "",
  pabda: "",
  gulsha: "",
  magur: "",
  carp: "",
  golda: "",
  bagda: "",
  tengra: "",
  feed_type: "",
  latitude: "",
  longitude: "",
  status: "Active",
  remarks: "",
};

const POULTRY_FIELDS = [
  "broiler",
  "layer",
  "sonali",
  "duck",
  "quail",
  "deshi_chicken",
];
const CATTLE_FIELDS = ["dairy_cow", "bull", "calf_heifer"];
const FISH_FIELDS = [
  "pangus",
  "tilapia",
  "koi",
  "shing",
  "pabda",
  "gulsha",
  "magur",
  "carp",
  "golda",
  "bagda",
  "tengra",
];

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

function getFarmerBadgeInfo(farmer: any) {
  if (farmer.farm_category === "Poultry") {
    if (Number(farmer.layer || 0) > 0 && Number(farmer.broiler || 0) === 0) {
      return { emoji: "🥚", label: "Layer", bg: "#d97706", border: "#b45309" };
    }
    if (Number(farmer.sonali || 0) > 0 && Number(farmer.broiler || 0) === 0) {
      return { emoji: "🐓", label: "Sonali", bg: "#ea580c", border: "#c2410c" };
    }
    if (Number(farmer.duck || 0) > 0 && Number(farmer.broiler || 0) === 0) {
      return { emoji: "🦆", label: "Duck", bg: "#0d9488", border: "#0f766e" };
    }
    return { emoji: "🐔", label: "Broiler", bg: "#f59e0b", border: "#d97706" };
  }
  if (farmer.farm_category === "Cattle") {
    return { emoji: "🐄", label: "Cattle", bg: "#16a34a", border: "#15803d" };
  }
  if (farmer.farm_category === "Fish") {
    return { emoji: "🐟", label: "Fish", bg: "#0284c7", border: "#0369a1" };
  }
  return { emoji: "🏡", label: "Farm", bg: "#475569", border: "#334155" };
}

// Clean Marker without circular background; auto-scales horizontal vs square/round logos
function buildCustomFarmMarker(farmer: any) {
  const { emoji, border } = getFarmerBadgeInfo(farmer);
  const logoUrl = farmer.companies?.logo_url;
  const safeName = (farmer.farmer_name || "Farmer")
    .replace(/</g, "&lt;")
    .slice(0, 16);

  const html = `
    <div style="display:flex;flex-direction:column;align-items:center;transform:translateY(-14px);">
      <div style="position:relative;display:flex;flex-direction:column;align-items:center;justify-content:center;">
        <span style="font-size:24px;line-height:1;filter:drop-shadow(0 2px 3px rgba(0,0,0,0.35));">${emoji}</span>
        ${
          logoUrl
            ? `<img src="${logoUrl}" onload="if(this.naturalWidth > this.naturalHeight * 1.3){this.style.width='44px';this.style.height='auto';this.style.maxHeight='20px';}else{this.style.width='22px';this.style.height='22px';}" style="margin-top:1px;width:24px;height:22px;object-fit:contain;background:rgba(255,255,255,0.92);padding:1px 2px;border-radius:3px;box-shadow:0 1px 2px rgba(0,0,0,0.2);" />`
            : ""
        }
      </div>
      <div style="margin-top:2px;background:rgba(255,255,255,0.95);border:1px solid ${border};color:#0f172a;font-size:10px;font-weight:700;padding:0px 5px;border-radius:4px;white-space:nowrap;box-shadow:0 1px 2px rgba(0,0,0,0.15);">
        ${safeName}
      </div>
    </div>
  `;

  return L.divIcon({
    className: "custom-farm-div-icon",
    html,
    iconSize: [52, 54],
    iconAnchor: [26, 44],
    popupAnchor: [0, -38],
  });
}

function getStockBreakdownList(farmer: any): { label: string; qty: number }[] {
  const list: { label: string; qty: number }[] = [];
  const allKeys = [...POULTRY_FIELDS, ...CATTLE_FIELDS, ...FISH_FIELDS];
  allKeys.forEach((k) => {
    const val = Number(farmer?.[k] || 0);
    if (val > 0) {
      const label = k
        .replace("_", " ")
        .replace(/\b\w/g, (c) => c.toUpperCase());
      list.push({ label, qty: val });
    }
  });
  return list;
}

function RecenterMap({ position }: { position: [number, number] }) {
  const map = useMap();
  useEffect(() => {
    if (!position) return;
    map.flyTo(position, 17, { animate: true, duration: 1.2 });
  }, [position, map]);
  return null;
}

function FitBounds({ positions }: { positions: [number, number][] }) {
  const map = useMap();
  useEffect(() => {
    if (positions.length === 0) return;
    map.fitBounds(positions, { padding: [45, 45] });
  }, [positions, map]);
  return null;
}

export default function FarmersPage() {
  const { role } = useAuth();
  const mapSectionRef = useRef<HTMLDivElement | null>(null);

  if (role === "viewer") {
    return <div className="p-6 text-center text-slate-500">Access Denied</div>;
  }

  const isFullAccess = role === "super_admin" || role === "manager";

  const [currentUserId, setCurrentUserId] = useState("");
  const [currentUserName, setCurrentUserName] = useState("User");
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>([]);

  const [farmers, setFarmers] = useState<any[]>([]);
  const [companies, setCompanies] = useState<any[]>([]);
  const [agents, setAgents] = useState<any[]>([]);

  const [search, setSearch] = useState("");
  const [editingId, setEditingId] = useState<number | null>(null);
  const [selectedFarmer, setSelectedFarmer] = useState<any>(null);
  const [showMap, setShowMap] = useState(false);
  const [showForm, setShowForm] = useState(false);
  const [showFarmerMap, setShowFarmerMap] = useState(false);
  const [showExportMenu, setShowExportMenu] = useState(false);
  const [showSummary, setShowSummary] = useState(false);
  const [form, setForm] = useState({ ...EMPTY_FORM });

  const [pickerCenterOverride, setPickerCenterOverride] = useState<
    [number, number] | null
  >(null);
  const [locatingInPicker, setLocatingInPicker] = useState(false);

  const [activeTrackingFarmerId, setActiveTrackingFarmerId] = useState<
    number | null
  >(null);
  const [farmerNotes, setFarmerNotes] = useState<any[]>([]);
  const [newNoteText, setNewNoteText] = useState("");
  const [loadingNotes, setLoadingNotes] = useState(false);
  const [savingNote, setSavingNote] = useState(false);

  const [companyFilter, setCompanyFilter] = useState("");
  const [agentFilter, setAgentFilter] = useState("");
  const [districtFilter, setDistrictFilter] = useState("");
  const [upazilaFilter, setUpazilaFilter] = useState("");
  const [unionFilter, setUnionFilter] = useState("");
  const [farmCategoryFilter, setFarmCategoryFilter] = useState("");

  const [districts, setDistricts] = useState<any[]>([]);
  const [upazilas, setUpazilas] = useState<any[]>([]);
  const [unions, setUnions] = useState<any[]>([]);
  const [filterUpazilas, setFilterUpazilas] = useState<any[]>([]);
  const [filterUnions, setFilterUnions] = useState<any[]>([]);

  const [myLocationMode, setMyLocationMode] = useState(false);
  const [selectedFarmerPosition, setSelectedFarmerPosition] = useState<
    [number, number] | null
  >(null);
  const [myPosition, setMyPosition] = useState<[number, number] | null>(null);
  const farmerMarkerRefs = useRef<Record<number, L.Marker | null>>({});

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
      farmerData,
      { data: companyData },
      { data: agentData },
      { data: districtData },
    ] = await Promise.all([
      getFarmers(),
      supabase.from("companies").select("*").order("company_name"),
      supabase.from("agents").select("*").order("agent_name"),
      supabase.from("districts").select("*").order("name"),
    ]);

    const allFarmers = farmerData || [];
    const allowedFarmers = isFullAccess
      ? allFarmers
      : allFarmers.filter((f: any) => {
          const matchUpa =
            uUpazilas.length === 0 || uUpazilas.includes(f.upazila);
          const matchDist =
            uDistricts.length === 0 || uDistricts.includes(f.district);
          return matchUpa && matchDist;
        });

    setFarmers(allowedFarmers);
    setCompanies(companyData || []);

    const allAgents = agentData || [];
    const allowedAgents = isFullAccess
      ? allAgents
      : allAgents.filter((a: any) => {
          const aDist = a.district || a.district_name || "";
          const aUpa = a.upazila || a.upazila_name || "";
          const matchUpa =
            uUpazilas.length === 0 || !aUpa || uUpazilas.includes(aUpa);
          const matchDist =
            uDistricts.length === 0 || !aDist || uDistricts.includes(aDist);
          const hasAllowedFarmer = allowedFarmers.some(
            (f: any) => String(f.agent_id) === String(a.id)
          );
          if (aUpa || aDist) return matchUpa && matchDist;
          return hasAllowedFarmer;
        });

    setAgents(allowedAgents);

    const allDistricts = districtData || [];
    const allowedDists = (
      isFullAccess
        ? allDistricts
        : allDistricts.filter(
            (d: any) => uDistricts.length === 0 || uDistricts.includes(d.name)
          )
    ).sort((a: any, b: any) => (a.name || "").localeCompare(b.name || ""));

    setDistricts(allowedDists);

    if (!districtFilter) {
      if (allowedDists.length === 1) {
        setDistrictFilter(allowedDists[0].name);
      } else if (allowedDists.length > 0 && !isFullAccess) {
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

  function clearFilters() {
    setCompanyFilter("");
    setAgentFilter("");
    setDistrictFilter(districts.length === 1 ? districts[0].name : "");
    setUpazilaFilter(
      !isFullAccess && assignedUpazilas.length === 1 ? assignedUpazilas[0] : ""
    );
    setUnionFilter("");
    setFarmCategoryFilter("");
    setSearch("");
  }

  async function loadFarmerNotes(farmerId: number) {
    setActiveTrackingFarmerId(farmerId);
    setLoadingNotes(true);
    try {
      const { data, error } = await supabase
        .from("farmer_trackings")
        .select("*")
        .eq("farmer_id", farmerId)
        .order("created_at", { ascending: true });

      if (!error) {
        setFarmerNotes(data || []);
      } else {
        setFarmerNotes([]);
      }
    } catch {
      setFarmerNotes([]);
    } finally {
      setLoadingNotes(false);
    }
  }

  useEffect(() => {
    if (selectedFarmer?.id) {
      setNewNoteText("");
      loadFarmerNotes(selectedFarmer.id);
    }
  }, [selectedFarmer]);

  async function handleAddFarmerNote(
    e: React.FormEvent,
    targetFarmerId?: number
  ) {
    e.preventDefault();
    const fId = targetFarmerId || selectedFarmer?.id || activeTrackingFarmerId;
    if (!fId || !newNoteText.trim()) return;

    setSavingNote(true);
    try {
      const { error } = await supabase.from("farmer_trackings").insert([
        {
          farmer_id: fId,
          note: newNoteText.trim(),
          user_id: currentUserId || null,
          user_name: currentUserName,
        },
      ]);

      if (error) throw error;
      setNewNoteText("");
      await loadFarmerNotes(fId);
    } catch (err: any) {
      alert(err?.message || "Failed to save tracking note.");
    } finally {
      setSavingNote(false);
    }
  }

  function handleViewFarmerOnMap(farmer: any) {
    if (!farmer.latitude || !farmer.longitude) {
      alert("No map coordinates saved for this farmer.");
      return;
    }
    const lat = Number(farmer.latitude);
    const lng = Number(farmer.longitude);
    setShowSummary(false);
    setShowFarmerMap(true);
    setMyLocationMode(false);
    setSelectedFarmerPosition([lat, lng]);
    loadFarmerNotes(farmer.id);

    setTimeout(() => {
      mapSectionRef.current?.scrollIntoView({
        behavior: "smooth",
        block: "center",
      });
      const marker = farmerMarkerRefs.current[farmer.id];
      if (marker) {
        marker.openPopup();
      }
    }, 350);
  }

  const selectedDistrict = districts.find((d) => d.name === form.district);

  useEffect(() => {
    async function loadUpazilas() {
      if (!selectedDistrict) {
        setUpazilas([]);
        return;
      }
      const { data } = await supabase
        .from("upazilas")
        .select("*")
        .eq("district_id", selectedDistrict.id)
        .order("name");

      const list = data || [];
      const allowed = isFullAccess
        ? list
        : list.filter(
            (u: any) =>
              assignedUpazilas.length === 0 || assignedUpazilas.includes(u.name)
          );
      setUpazilas(allowed);

      // Auto-select first upazila if none selected or not in list
      if (
        allowed.length > 0 &&
        (!form.upazila || !allowed.some((u: any) => u.name === form.upazila))
      ) {
        setForm((prev) => ({ ...prev, upazila: allowed[0].name }));
      }
    }
    loadUpazilas();
  }, [selectedDistrict, isFullAccess, assignedUpazilas]);

  const selectedUpazila = upazilas.find((u) => u.name === form.upazila);

  useEffect(() => {
    async function loadUnions() {
      if (!selectedUpazila) {
        setUnions([]);
        return;
      }
      const { data } = await supabase
        .from("unions")
        .select("*")
        .eq("upazila_id", selectedUpazila.id)
        .order("name");
      setUnions(data || []);
    }
    loadUnions();
  }, [selectedUpazila]);

  // Form Agents filtered by selected Company AND selected Upazila
  const formAvailableAgents = agents.filter((a: any) => {
    if (form.company_id && String(a.company_id) !== String(form.company_id)) {
      return false;
    }
    const aUpa = a.upazila || a.upazila_name || "";
    if (form.upazila && aUpa && aUpa !== form.upazila) {
      return false;
    }
    return true;
  });

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
  }, [upazilaFilter, districtFilter, filterUpazilas]);

  const availableFilterAgents = agents.filter((a: any) => {
    if (companyFilter && String(a.company_id) !== String(companyFilter)) {
      return false;
    }
    const aDist = a.district || a.district_name || "";
    const aUpa = a.upazila || a.upazila_name || "";

    if (districtFilter) {
      if (aDist && aDist !== districtFilter) return false;
      if (!aDist) {
        const hasFarmerInDist = farmers.some(
          (f) =>
            String(f.agent_id) === String(a.id) && f.district === districtFilter
        );
        if (!hasFarmerInDist) return false;
      }
    }

    if (upazilaFilter) {
      if (aUpa && aUpa !== upazilaFilter) return false;
      if (!aUpa) {
        const hasFarmerInUpa = farmers.some(
          (f) =>
            String(f.agent_id) === String(a.id) && f.upazila === upazilaFilter
        );
        if (!hasFarmerInUpa) return false;
      }
    }

    return true;
  });

  async function handleSave() {
    if (
      !form.company_id ||
      !form.agent_id ||
      !form.farmer_name.trim() ||
      !form.farm_category ||
      !form.district ||
      !form.upazila
    ) {
      alert(
        "Company, District, Upazila, Agent, Farmer Name and Farm Category are required"
      );
      return;
    }

    const payload: any = {
      ...form,
      company_id: Number(form.company_id),
      agent_id: Number(form.agent_id),
      latitude: form.latitude === "" ? null : Number(form.latitude),
      longitude: form.longitude === "" ? null : Number(form.longitude),
    };

    [...POULTRY_FIELDS, ...CATTLE_FIELDS, ...FISH_FIELDS].forEach((field) => {
      payload[field] = Number((form as any)[field] || 0);
    });

    if (editingId) {
      await updateFarmer(editingId, payload);
      alert("Farmer Updated");
    } else {
      await createFarmer(payload);
      alert("Farmer Saved");
    }

    setEditingId(null);
    setForm({ ...EMPTY_FORM });
    setShowForm(false);
    loadData();
  }

  async function handleDelete(id: number) {
    if (!confirm("Delete Farmer?")) return;
    await deleteFarmer(id);
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

  const filteredFarmers = farmers
    .filter((f) => {
      const q = search.toLowerCase();
      const matchSearch =
        !q ||
        f.farmer_name?.toLowerCase().includes(q) ||
        f.mobile?.toLowerCase().includes(q) ||
        f.village?.toLowerCase().includes(q);
      const matchCompany =
        companyFilter === "" || String(f.company_id) === companyFilter;
      const matchAgent =
        agentFilter === "" || String(f.agent_id) === agentFilter;
      const matchDistrict =
        districtFilter === "" || f.district === districtFilter;
      const matchUpazila =
        upazilaFilter === "" || f.upazila === upazilaFilter;
      const matchUnion = unionFilter === "" || f.union_name === unionFilter;
      const matchCategory =
        farmCategoryFilter === "" || f.farm_category === farmCategoryFilter;

      return (
        matchSearch &&
        matchCompany &&
        matchAgent &&
        matchDistrict &&
        matchUpazila &&
        matchUnion &&
        matchCategory
      );
    })
    .sort((a: any, b: any) => {
      const agentCompare = (a.agents?.agent_name || "").localeCompare(
        b.agents?.agent_name || ""
      );
      if (agentCompare !== 0) return agentCompare;
      return (a.farmer_name || "").localeCompare(b.farmer_name || "");
    });

  const agentGroupedFarmers = Object.values(
    filteredFarmers.reduce((acc: Record<string, any>, farmer: any) => {
      const key = String(farmer.agent_id || "unassigned");
      if (!acc[key]) {
        acc[key] = {
          agentId: key,
          agentName: farmer.agents?.agent_name || "Unassigned Agent",
          companyName: farmer.companies?.company_name || "-",
          upazila: farmer.upazila || "",
          farmers: [],
        };
      }
      acc[key].farmers.push(farmer);
      return acc;
    }, {})
  );

  const nearestFarmers = myPosition
    ? filteredFarmers
        .filter(
          (f) =>
            f.latitude &&
            f.longitude &&
            !isNaN(Number(f.latitude)) &&
            !isNaN(Number(f.longitude))
        )
        .map((f) => ({
          ...f,
          distance: Math.round(
            Math.sqrt(
              Math.pow(Number(f.latitude) - myPosition[0], 2) +
                Math.pow(Number(f.longitude) - myPosition[1], 2)
            ) * 111
          ),
        }))
        .sort((a, b) => a.distance - b.distance)
        .slice(0, 10)
    : [];

  const companySummary = Object.values(
    filteredFarmers.reduce((acc: any, farmer: any) => {
      const company = farmer.companies?.company_name || "Unknown";
      if (!acc[company]) {
        acc[company] = {
          company,
          totalFarmers: 0,
          poultryFarmers: 0,
          cattleFarmers: 0,
          fishFarmers: 0,
          broiler: 0,
          layer: 0,
          sonali: 0,
          duck: 0,
          quail: 0,
          deshiChicken: 0,
          dairyCow: 0,
          bull: 0,
          calfHeifer: 0,
          pangus: 0,
          tilapia: 0,
          koi: 0,
          shing: 0,
          pabda: 0,
          gulsha: 0,
          magur: 0,
          carp: 0,
          golda: 0,
          bagda: 0,
          tengra: 0,
          totalPoultry: 0,
          totalCattle: 0,
          totalFish: 0,
          broilerFarmers: 0,
          layerFarmers: 0,
          sonaliFarmers: 0,
          duckFarmers: 0,
          quailFarmers: 0,
          deshiFarmers: 0,
          dairyCowFarmers: 0,
          bullFarmers: 0,
          calfHeiferFarmers: 0,
          pangusFarmers: 0,
          tilapiaFarmers: 0,
          koiFarmers: 0,
          shingFarmers: 0,
          pabdaFarmers: 0,
          gulshaFarmers: 0,
          magurFarmers: 0,
          carpFarmers: 0,
          goldaFarmers: 0,
          bagdaFarmers: 0,
          tengraFarmers: 0,
        };
      }

      const c = acc[company];
      c.totalFarmers++;

      if (Number(farmer.broiler || 0) > 0) c.broilerFarmers++;
      if (Number(farmer.layer || 0) > 0) c.layerFarmers++;
      if (Number(farmer.sonali || 0) > 0) c.sonaliFarmers++;
      if (Number(farmer.duck || 0) > 0) c.duckFarmers++;
      if (Number(farmer.quail || 0) > 0) c.quailFarmers++;
      if (Number(farmer.deshi_chicken || 0) > 0) c.deshiFarmers++;

      if (farmer.farm_category === "Poultry") {
        c.poultryFarmers++;
        c.broiler += Number(farmer.broiler || 0);
        c.layer += Number(farmer.layer || 0);
        c.sonali += Number(farmer.sonali || 0);
        c.duck += Number(farmer.duck || 0);
        c.quail += Number(farmer.quail || 0);
        c.deshiChicken += Number(farmer.deshi_chicken || 0);
        c.totalPoultry =
          c.broiler + c.layer + c.sonali + c.duck + c.quail + c.deshiChicken;
      }

      if (Number(farmer.dairy_cow || 0) > 0) c.dairyCowFarmers++;
      if (Number(farmer.bull || 0) > 0) c.bullFarmers++;
      if (Number(farmer.calf_heifer || 0) > 0) c.calfHeiferFarmers++;

      if (farmer.farm_category === "Cattle") {
        c.cattleFarmers++;
        c.dairyCow += Number(farmer.dairy_cow || 0);
        c.bull += Number(farmer.bull || 0);
        c.calfHeifer += Number(farmer.calf_heifer || 0);
        c.totalCattle = c.dairyCow + c.bull + c.calfHeifer;
      }

      if (farmer.farm_category === "Fish") {
        c.fishFarmers++;
        FISH_FIELDS.forEach((f) => {
          const val = Number(farmer[f] || 0);
          c[f] += val;
          if (val > 0) c[`${f}Farmers`]++;
        });
        c.totalFish = FISH_FIELDS.reduce((sum, f) => sum + c[f], 0);
      }

      return acc;
    }, {})
  );

  const selectedCompanyName =
    companies.find((c) => String(c.id) === companyFilter)?.company_name ||
    "All Companies";

  function styleStandardWorksheet(
    worksheet: ExcelJS.Worksheet,
    headerColor: string,
    widths: number[]
  ) {
    worksheet.getRow(1).eachCell((cell) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: headerColor },
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
    });

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber > 1) {
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

    widths.forEach((w, i) => {
      worksheet.getColumn(i + 1).width = w;
    });
    worksheet.getRow(1).height = 28;
    worksheet.views = [{ state: "frozen", ySplit: 1 }];
    worksheet.pageSetup.orientation = "landscape";
    worksheet.pageSetup.fitToPage = true;
    worksheet.pageSetup.fitToWidth = 1;
    worksheet.pageSetup.fitToHeight = 0;
  }

  async function saveWorkbook(workbook: ExcelJS.Workbook, filename: string) {
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(
      new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }),
      filename
    );
  }

  async function exportAllFarmersExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("All Farmers");
      worksheet.views = [{ state: "frozen", ySplit: 2 }];

      const generalHeaders = [
        "Company",
        "Agent",
        "Farmer",
        "Mobile",
        "Village",
        "Union",
        "Upazila",
        "District",
        "Farm Category",
      ];
      generalHeaders.forEach((h, idx) => {
        const col = idx + 1;
        worksheet.mergeCells(1, col, 2, col);
        worksheet.getCell(1, col).value = h;
      });

      worksheet.mergeCells("J1:O1");
      worksheet.getCell("J1").value = "POULTRY";
      ["Broiler", "Layer", "Sonali", "Duck", "Quail", "Deshi Chicken"].forEach(
        (v, i) => (worksheet.getCell(2, 10 + i).value = v)
      );

      worksheet.mergeCells("P1:R1");
      worksheet.getCell("P1").value = "CATTLE";
      ["Dairy Cow", "Bull", "Calf / Heifer"].forEach(
        (v, i) => (worksheet.getCell(2, 16 + i).value = v)
      );

      worksheet.mergeCells("S1:AC1");
      worksheet.getCell("S1").value = "FISH";
      [
        "Pangus",
        "Tilapia",
        "Koi",
        "Shing",
        "Pabda",
        "Gulsha",
        "Magur",
        "Carp",
        "Golda",
        "Bagda",
        "Tengra",
      ].forEach((v, i) => (worksheet.getCell(2, 19 + i).value = v));

      worksheet.mergeCells("AD1:AD2");
      worksheet.getCell("AD1").value = "Feed Type";
      worksheet.mergeCells("AE1:AE2");
      worksheet.getCell("AE1").value = "Status";

      for (let col = 1; col <= 31; col++) {
        if (col <= 9 || col >= 30) {
          const c = worksheet.getCell(1, col);
          c.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "1E3A8A" },
          };
          c.font = { bold: true, color: { argb: "FFFFFFFF" } };
          c.alignment = {
            horizontal: "center",
            vertical: "middle",
            wrapText: true,
          };
        }
      }

      const styleGroup = (
        mainCell: string,
        mainBg: string,
        startCol: number,
        endCol: number,
        subBg: string,
        subFg: string
      ) => {
        const mc = worksheet.getCell(mainCell);
        mc.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: mainBg },
        };
        mc.font = { bold: true, color: { argb: "FFFFFFFF" } };
        mc.alignment = { horizontal: "center", vertical: "middle" };
        for (let col = startCol; col <= endCol; col++) {
          const sc = worksheet.getCell(2, col);
          sc.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: subBg },
          };
          sc.font = { bold: true, color: { argb: subFg } };
          sc.alignment = {
            horizontal: "center",
            vertical: "middle",
            wrapText: true,
          };
        }
      };

      styleGroup("J1", "D97706", 10, 15, "FEF3C7", "92400E");
      styleGroup("P1", "15803D", 16, 18, "DCFCE7", "166534");
      styleGroup("S1", "0284C7", 19, 29, "E0F2FE", "075985");

      for (let r = 1; r <= 2; r++) {
        for (let c = 1; c <= 31; c++) {
          worksheet.getCell(r, c).border = {
            top: { style: "thin", color: { argb: "CBD5E1" } },
            left: { style: "thin", color: { argb: "CBD5E1" } },
            bottom: { style: "thin", color: { argb: "CBD5E1" } },
            right: { style: "thin", color: { argb: "CBD5E1" } },
          };
        }
      }

      filteredFarmers.forEach((f: any) => {
        const row = worksheet.addRow([
          f.companies?.company_name || "",
          f.agents?.agent_name || "",
          f.farmer_name || "",
          f.mobile || "",
          f.village || "",
          f.union_name || "",
          f.upazila || "",
          f.district || "",
          f.farm_category || "",
          ...POULTRY_FIELDS.map((k) => Number(f[k] || 0)),
          ...CATTLE_FIELDS.map((k) => Number(f[k] || 0)),
          ...FISH_FIELDS.map((k) => Number(f[k] || 0)),
          f.feed_type || "",
          f.status || "",
        ]);
        row.eachCell((cell) => {
          cell.border = {
            top: { style: "thin", color: { argb: "D1D5DB" } },
            left: { style: "thin", color: { argb: "D1D5DB" } },
            bottom: { style: "thin", color: { argb: "D1D5DB" } },
            right: { style: "thin", color: { argb: "D1D5DB" } },
          };
          cell.alignment = { vertical: "middle", wrapText: true };
        });
      });

      worksheet.getRow(1).height = 26;
      worksheet.getRow(2).height = 30;
      const widths = [
        22, 20, 22, 15, 18, 18, 18, 18, 18, 12, 12, 12, 12, 12, 16, 14, 12, 16,
        12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 18, 12,
      ];
      widths.forEach((w, i) => (worksheet.getColumn(i + 1).width = w));
      worksheet.autoFilter = { from: "A2", to: "AE2" };

      await saveWorkbook(
        workbook,
        `All_Farmers_Report_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch {
      alert("Excel export failed.");
    }
  }

  async function exportPoultryExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Poultry Farmers");
      const list = filteredFarmers.filter((f) => f.farm_category === "Poultry");

      worksheet.addRow([
        "Company",
        "Agent",
        "Farmer",
        "Mobile",
        "Village",
        "Union",
        "Upazila",
        "District",
        "Farm Category",
        "Broiler",
        "Layer",
        "Sonali",
        "Duck",
        "Quail",
        "Deshi Chicken",
        "Feed Type",
        "Latitude",
        "Longitude",
        "Status",
        "Remarks",
      ]);

      list.forEach((f) => {
        worksheet.addRow([
          f.companies?.company_name || "",
          f.agents?.agent_name || "",
          f.farmer_name || "",
          f.mobile || "",
          f.village || "",
          f.union_name || "",
          f.upazila || "",
          f.district || "",
          f.farm_category || "",
          ...POULTRY_FIELDS.map((k) => Number(f[k] || 0)),
          f.feed_type || "",
          f.latitude ?? "",
          f.longitude ?? "",
          f.status || "",
          f.remarks || "",
        ]);
      });

      styleStandardWorksheet(
        worksheet,
        "D97706",
        [
          25, 22, 25, 16, 20, 20, 18, 18, 16, 12, 12, 12, 12, 12, 18, 18, 14,
          14, 12, 30,
        ]
      );
      await saveWorkbook(
        workbook,
        `Poultry_Farmers_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch {
      alert("Poultry export failed.");
    }
  }

  async function exportCattleExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Cattle Farmers");
      const list = filteredFarmers.filter((f) => f.farm_category === "Cattle");

      worksheet.addRow([
        "Company",
        "Agent",
        "Farmer",
        "Mobile",
        "Village",
        "Union",
        "Upazila",
        "District",
        "Farm Category",
        "Dairy Cow",
        "Bull",
        "Calf / Heifer",
        "Total Cattle",
        "Feed Type",
        "Latitude",
        "Longitude",
        "Status",
        "Remarks",
      ]);

      list.forEach((f) => {
        const dc = Number(f.dairy_cow || 0);
        const bl = Number(f.bull || 0);
        const ch = Number(f.calf_heifer || 0);
        worksheet.addRow([
          f.companies?.company_name || "",
          f.agents?.agent_name || "",
          f.farmer_name || "",
          f.mobile || "",
          f.village || "",
          f.union_name || "",
          f.upazila || "",
          f.district || "",
          f.farm_category || "",
          dc,
          bl,
          ch,
          dc + bl + ch,
          f.feed_type || "",
          f.latitude ?? "",
          f.longitude ?? "",
          f.status || "",
          f.remarks || "",
        ]);
      });

      styleStandardWorksheet(
        worksheet,
        "15803D",
        [
          25, 22, 25, 16, 20, 20, 18, 18, 16, 15, 12, 18, 18, 18, 14, 14, 12,
          30,
        ]
      );
      await saveWorkbook(
        workbook,
        `Cattle_Farmers_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch {
      alert("Cattle export failed.");
    }
  }

  async function exportFishExcel() {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Fish Farmers");
      const list = filteredFarmers.filter((f) => f.farm_category === "Fish");

      worksheet.addRow([
        "Company",
        "Agent",
        "Farmer",
        "Mobile",
        "Village",
        "Union",
        "Upazila",
        "District",
        "Farm Category",
        "Pangus",
        "Tilapia",
        "Koi",
        "Shing",
        "Pabda",
        "Gulsha",
        "Magur",
        "Carp",
        "Golda",
        "Bagda",
        "Tengra",
        "Total Fish",
        "Feed Type",
        "Latitude",
        "Longitude",
        "Status",
        "Remarks",
      ]);

      list.forEach((f) => {
        const vals = FISH_FIELDS.map((k) => Number(f[k] || 0));
        const total = vals.reduce((a, b) => a + b, 0);
        worksheet.addRow([
          f.companies?.company_name || "",
          f.agents?.agent_name || "",
          f.farmer_name || "",
          f.mobile || "",
          f.village || "",
          f.union_name || "",
          f.upazila || "",
          f.district || "",
          f.farm_category || "",
          ...vals,
          total,
          f.feed_type || "",
          f.latitude ?? "",
          f.longitude ?? "",
          f.status || "",
          f.remarks || "",
        ]);
      });

      styleStandardWorksheet(
        worksheet,
        "0284C7",
        [
          25, 22, 25, 16, 20, 20, 18, 18, 16, 12, 12, 12, 12, 12, 12, 12, 12,
          12, 12, 12, 15, 18, 14, 14, 12, 30,
        ]
      );
      await saveWorkbook(
        workbook,
        `Fish_Farmers_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch {
      alert("Fish export failed.");
    }
  }

  function exportFarmerSummaryExcel() {
    try {
      const exportData = companySummary.map((row: any) => ({
        Company: row.company,
        Total_Farmers: row.totalFarmers,
        Poultry_Farmers: row.poultryFarmers,
        Cattle_Farmers: row.cattleFarmers,
        Fish_Farmers: row.fishFarmers,
      }));
      const worksheet = XLSX.utils.json_to_sheet(exportData);
      worksheet["!cols"] = [
        { wch: 28 },
        { wch: 18 },
        { wch: 20 },
        { wch: 20 },
        { wch: 18 },
      ];
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Company Summary");
      const excelBuffer = XLSX.write(workbook, {
        bookType: "xlsx",
        type: "array",
      });
      const today = new Date().toISOString().split("T")[0];
      saveAs(
        new Blob([excelBuffer], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
        `Company_Farmer_Summary_${today}.xlsx`
      );
    } catch {
      alert("Company Summary export failed.");
    }
  }

  const exportPoultrySummaryExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Poultry Summary");
      worksheet.addRow([
        "Company",
        "Broiler Farms",
        "Total Broiler",
        "Layer Farms",
        "Total Layer",
        "Sonali Farms",
        "Total Sonali",
        "Duck Farms",
        "Total Duck",
        "Quail Farms",
        "Total Quail",
        "Deshi Chicken Farms",
        "Total Deshi Chicken",
        "Total Poultry Farms",
        "Total Poultry",
      ]);

      companySummary.forEach((row: any) => {
        worksheet.addRow([
          row.company,
          row.broilerFarmers,
          row.broiler,
          row.layerFarmers,
          row.layer,
          row.sonaliFarmers,
          row.sonali,
          row.duckFarmers,
          row.duck,
          row.quailFarmers,
          row.quail,
          row.deshiFarmers,
          row.deshiChicken,
          row.poultryFarmers,
          row.totalPoultry,
        ]);
      });

      styleStandardWorksheet(
        worksheet,
        "D97706",
        [25, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18]
      );
      await saveWorkbook(
        workbook,
        `Poultry_Summary_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch {
      alert("Poultry summary export failed.");
    }
  };

  const exportCattleSummaryExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Cattle Summary");
      worksheet.addRow([
        "Company",
        "Dairy Farms",
        "Total Dairy Cow",
        "Bull Farms",
        "Total Bull",
        "Calf / Heifer Farms",
        "Total Calf / Heifer",
        "Total Cattle",
      ]);
      companySummary.forEach((row: any) => {
        worksheet.addRow([
          row.company,
          row.dairyCowFarmers,
          row.dairyCow,
          row.bullFarmers,
          row.bull,
          row.calfHeiferFarmers,
          row.calfHeifer,
          row.totalCattle,
        ]);
      });
      styleStandardWorksheet(
        worksheet,
        "15803D",
        [28, 16, 16, 16, 16, 18, 18, 18]
      );
      await saveWorkbook(
        workbook,
        `Cattle_Summary_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch {
      alert("Cattle Summary export failed.");
    }
  };

  const exportFishSummaryExcel = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Fish Summary");
      worksheet.addRow([
        "Company",
        "Pangus Farms",
        "Total Pangus",
        "Tilapia Farms",
        "Total Tilapia",
        "Koi Farms",
        "Total Koi",
        "Shing Farms",
        "Total Shing",
        "Pabda Farms",
        "Total Pabda",
        "Gulsha Farms",
        "Total Gulsha",
        "Magur Farms",
        "Total Magur",
        "Carp Farms",
        "Total Carp",
        "Golda Farms",
        "Total Golda",
        "Bagda Farms",
        "Total Bagda",
        "Tengra Farms",
        "Total Tengra",
        "Total Fish Farms",
        "Total Fish",
      ]);

      companySummary.forEach((row: any) => {
        worksheet.addRow([
          row.company,
          row.pangusFarmers,
          row.pangus,
          row.tilapiaFarmers,
          row.tilapia,
          row.koiFarmers,
          row.koi,
          row.shingFarmers,
          row.shing,
          row.pabdaFarmers,
          row.pabda,
          row.gulshaFarmers,
          row.gulsha,
          row.magurFarmers,
          row.magur,
          row.carpFarmers,
          row.carp,
          row.goldaFarmers,
          row.golda,
          row.bagdaFarmers,
          row.bagda,
          row.tengraFarmers,
          row.tengra,
          row.fishFarmers,
          row.totalFish,
        ]);
      });

      styleStandardWorksheet(
        worksheet,
        "0284C7",
        [25, ...Array(24).fill(16)]
      );
      await saveWorkbook(
        workbook,
        `Fish_Summary_${new Date().toISOString().slice(0, 10)}.xlsx`
      );
    } catch {
      alert("Fish summary export failed.");
    }
  };

  function openEditForm(farmer: any) {
    setEditingId(farmer.id);
    setForm({
      company_id: String(farmer.company_id),
      agent_id: String(farmer.agent_id),
      farmer_name: farmer.farmer_name || "",
      mobile: farmer.mobile || "",
      village: farmer.village || "",
      union_name: farmer.union_name || "",
      upazila: farmer.upazila || "",
      district: farmer.district || "",
      farm_category: farmer.farm_category || "",
      broiler: String(farmer.broiler || ""),
      layer: String(farmer.layer || ""),
      sonali: String(farmer.sonali || ""),
      duck: String(farmer.duck || ""),
      quail: String(farmer.quail || ""),
      deshi_chicken: String(farmer.deshi_chicken || ""),
      dairy_cow: String(farmer.dairy_cow || ""),
      bull: String(farmer.bull || ""),
      calf_heifer: String(farmer.calf_heifer || ""),
      pangus: String(farmer.pangus || ""),
      tilapia: String(farmer.tilapia || ""),
      koi: String(farmer.koi || ""),
      shing: String(farmer.shing || ""),
      pabda: String(farmer.pabda || ""),
      gulsha: String(farmer.gulsha || ""),
      magur: String(farmer.magur || ""),
      carp: String(farmer.carp || ""),
      golda: String(farmer.golda || ""),
      bagda: String(farmer.bagda || ""),
      tengra: String(farmer.tengra || ""),
      feed_type: farmer.feed_type || "",
      latitude: String(farmer.latitude || ""),
      longitude: String(farmer.longitude || ""),
      status: farmer.status || "Active",
      remarks: farmer.remarks || "",
    });
    setShowForm(true);
  }

  return (
    <div className="p-2.5 sm:p-4 max-w-7xl mx-auto space-y-3">
      {/* 1. COMPACT FILTER BAR */}
      <div className="bg-white border border-slate-300 rounded-xl p-2.5 shadow-sm">
        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-2 items-center">
          <div className="col-span-2 sm:col-span-2 lg:col-span-2 relative">
            <Search
              size={13}
              className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400"
            />
            <input
              type="text"
              placeholder="Search farmer, mobile, village..."
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
              setAgentFilter("");
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
              setAgentFilter("");
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
            onChange={(e) => setUnionFilter(e.target.value)}
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
            value={farmCategoryFilter}
            onChange={(e) => setFarmCategoryFilter(e.target.value)}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Categories</option>
            <option value="Poultry">🐔 Poultry</option>
            <option value="Cattle">🐄 Cattle</option>
            <option value="Fish">🐟 Fish</option>
          </select>

          <select
            value={companyFilter}
            onChange={(e) => {
              setCompanyFilter(e.target.value);
              setAgentFilter("");
            }}
            className="rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
          >
            <option value="">All Companies</option>
            {companies.map((c) => (
              <option key={c.id} value={c.id}>
                {c.company_name}
              </option>
            ))}
          </select>

          <div className="flex items-center gap-1.5">
            <select
              value={agentFilter}
              onChange={(e) => setAgentFilter(e.target.value)}
              className="w-full rounded-lg border border-slate-300 bg-white px-2 py-1.5 text-xs outline-none focus:border-blue-500"
            >
              <option value="">All Agents</option>
              {availableFilterAgents.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.agent_name}
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

      {/* 2. ACTION BAR */}
      <div className="flex flex-wrap items-center justify-between gap-2 bg-white border border-slate-300 rounded-xl px-3.5 py-2 shadow-sm">
        <div className="flex flex-wrap items-center gap-2">
          <button
            onClick={() => {
              setEditingId(null);
              const nourishComp =
                companies.find((c) =>
                  c.company_name?.toLowerCase().includes("nourish")
                ) || companies[0];
              const defCompId = nourishComp ? String(nourishComp.id) : "";

              // Chronologically sorted assigned districts: default to 1st
              const sortedDists = [...districts].sort((a, b) =>
                (a.name || "").localeCompare(b.name || "")
              );
              const defDist = sortedDists.length > 0 ? sortedDists[0].name : "";
              const defUpa =
                !isFullAccess && assignedUpazilas.length > 0
                  ? [...assignedUpazilas].sort()[0]
                  : "";

              setForm({
                ...EMPTY_FORM,
                company_id: defCompId,
                district: defDist,
                upazila: defUpa,
              });
              setShowForm(true);
            }}
            className="inline-flex items-center gap-1.5 bg-green-600 hover:bg-green-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <Plus size={14} /> Add Farmer
          </button>

          <button
            onClick={() => {
              setShowFarmerMap(false);
              setShowSummary(true);
            }}
            className="inline-flex items-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
          >
            <BarChart3 size={14} /> Summary
          </button>

          <div className="relative">
            <button
              onClick={() => setShowExportMenu(!showExportMenu)}
              className="inline-flex items-center gap-1.5 bg-emerald-700 hover:bg-emerald-800 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
            >
              <FileSpreadsheet size={14} /> Export Excel
              <span className="text-[10px]">{showExportMenu ? "▲" : "▼"}</span>
            </button>

            {showExportMenu && (
              <div className="absolute left-0 mt-1.5 w-48 bg-white border border-slate-200 rounded-xl shadow-xl z-[2000] text-xs overflow-hidden">
                <button
                  onClick={() => {
                    exportAllFarmersExcel();
                    setShowExportMenu(false);
                  }}
                  className="w-full text-left px-3.5 py-2.5 hover:bg-slate-50 border-b font-medium"
                >
                  📄 Export All Farmers
                </button>
                <button
                  onClick={() => {
                    exportPoultryExcel();
                    setShowExportMenu(false);
                  }}
                  className="w-full text-left px-3.5 py-2.5 hover:bg-amber-50 border-b font-medium"
                >
                  🐔 Export Poultry Farmers
                </button>
                <button
                  onClick={() => {
                    exportCattleExcel();
                    setShowExportMenu(false);
                  }}
                  className="w-full text-left px-3.5 py-2.5 hover:bg-emerald-50 border-b font-medium"
                >
                  🐄 Export Cattle Farmers
                </button>
                <button
                  onClick={() => {
                    exportFishExcel();
                    setShowExportMenu(false);
                  }}
                  className="w-full text-left px-3.5 py-2.5 hover:bg-sky-50 font-medium"
                >
                  🐟 Export Fish Farmers
                </button>
              </div>
            )}
          </div>

          <button
            onClick={() => {
              if (!showFarmerMap) {
                setMyLocationMode(false);
                setMyPosition(null);
                setSelectedFarmerPosition(null);
              }
              setShowSummary(false);
              setShowFarmerMap(!showFarmerMap);
            }}
            className={`inline-flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition ${
              showFarmerMap
                ? "bg-red-600 hover:bg-red-700 text-white"
                : "bg-blue-600 hover:bg-blue-700 text-white"
            }`}
          >
            <MapIcon size={14} />
            {showFarmerMap ? "Close Map" : "View Map"}
          </button>
        </div>

        <div className="text-xs font-semibold text-slate-600">
          Total Farmers:{" "}
          <span className="text-slate-900">{filteredFarmers.length}</span>
        </div>
      </div>

      {/* 3. INTERACTIVE MAP VIEW WITH FLOATING CLOSE BUTTON */}
      {showFarmerMap && (
        <div
          ref={mapSectionRef}
          className="bg-white rounded-xl shadow border border-slate-300 p-3 relative z-10"
        >
          <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
            <h3 className="text-sm font-bold text-slate-800">Farmer Map</h3>
            <div className="flex flex-wrap items-center gap-2 text-[11px] font-semibold">
              <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800">
                🐔 Broiler
              </span>
              <span className="px-2 py-0.5 rounded bg-orange-100 text-orange-800">
                🥚 Layer
              </span>
              <span className="px-2 py-0.5 rounded bg-red-100 text-red-800">
                🐓 Sonali
              </span>
              <span className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800">
                🐄 Cattle
              </span>
              <span className="px-2 py-0.5 rounded bg-sky-100 text-sky-800">
                🐟 Fish
              </span>
            </div>
          </div>

          <div className="w-full overflow-hidden rounded-xl border border-slate-300 relative">
            <MapContainer
              center={[24.7471, 90.4203]}
              zoom={8}
              style={{ height: "480px", width: "100%", zIndex: 1 }}
            >
              <MapLayers />

              {myLocationMode && myPosition ? (
                <RecenterMap position={myPosition} />
              ) : selectedFarmerPosition ? (
                <RecenterMap position={selectedFarmerPosition} />
              ) : null}

              {myPosition && (
                <Marker position={myPosition}>
                  <Tooltip permanent direction="top" offset={[0, -15]}>
                    📍 My Location
                  </Tooltip>
                  <Popup>
                    <div className="max-h-60 overflow-y-auto w-52 text-xs space-y-1.5">
                      <h3 className="font-bold text-slate-800">
                        📍 Nearest Farmers
                      </h3>
                      {nearestFarmers.length === 0 ? (
                        <div>No nearby farmers found.</div>
                      ) : (
                        nearestFarmers.map((farmer, index) => (
                          <div key={farmer.id} className="border-b pb-1">
                            <button
                              onClick={() => {
                                setSelectedFarmerPosition([
                                  Number(farmer.latitude),
                                  Number(farmer.longitude),
                                ]);
                                setMyLocationMode(false);
                              }}
                              className="text-blue-600 font-bold hover:underline text-left"
                            >
                              {index + 1}. {farmer.farmer_name}
                            </button>
                            <div className="text-slate-500">
                              {farmer.distance} km away
                            </div>
                          </div>
                        ))
                      )}
                    </div>
                  </Popup>
                </Marker>
              )}

              {!selectedFarmerPosition && (
                <FitBounds
                  positions={
                    myLocationMode
                      ? []
                      : filteredFarmers
                          .filter((f) => f.latitude && f.longitude)
                          .map((f) => [
                            Number(f.latitude),
                            Number(f.longitude),
                          ])
                  }
                />
              )}

              {filteredFarmers
                .filter((f) => f.latitude && f.longitude)
                .map((farmer) => {
                  const badge = getFarmerBadgeInfo(farmer);
                  const breakdown = getStockBreakdownList(farmer);

                  return (
                    <Marker
                      key={farmer.id}
                      ref={(ref) => {
                        farmerMarkerRefs.current[farmer.id] = ref;
                      }}
                      position={[
                        Number(farmer.latitude),
                        Number(farmer.longitude),
                      ]}
                      icon={buildCustomFarmMarker(farmer)}
                      eventHandlers={{
                        click: () => {
                          setNewNoteText("");
                          loadFarmerNotes(farmer.id);
                        },
                      }}
                    >
                      <Popup maxWidth={340}>
                        <div className="w-[270px] sm:w-[290px] text-xs space-y-2">
                          <div className="flex items-center gap-2 border-b pb-1.5">
                            {farmer.companies?.logo_url && (
                              <img
                                src={farmer.companies?.logo_url}
                                className="w-9 h-9 object-contain border rounded p-0.5 bg-white shrink-0"
                              />
                            )}
                            <div>
                              <h3 className="font-bold text-sm text-slate-800 leading-tight">
                                {badge.emoji} {farmer.farmer_name}
                              </h3>
                              <p className="text-[11px] text-slate-500">
                                {farmer.companies?.company_name} •{" "}
                                {farmer.agents?.agent_name}
                              </p>
                            </div>
                          </div>

                          <div className="bg-emerald-50 border border-emerald-200 rounded-lg p-2 space-y-1.5">
                            <div className="font-bold text-[11px] text-emerald-900">
                              📝 Farmer Tracking
                            </div>
                            <div className="max-h-24 overflow-y-auto space-y-1 pr-0.5">
                              {loadingNotes &&
                              activeTrackingFarmerId === farmer.id ? (
                                <div className="text-[11px] text-slate-400">
                                  Loading...
                                </div>
                              ) : activeTrackingFarmerId === farmer.id &&
                                farmerNotes.length > 0 ? (
                                farmerNotes.map((n) => (
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
                              onSubmit={(e) =>
                                handleAddFarmerNote(e, farmer.id)
                              }
                              className="flex gap-1"
                            >
                              <input
                                type="text"
                                placeholder="Write note..."
                                value={
                                  activeTrackingFarmerId === farmer.id
                                    ? newNoteText
                                    : ""
                                }
                                onChange={(e) => {
                                  setActiveTrackingFarmerId(farmer.id);
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
                                  {farmer.mobile || "-"}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Address
                                </td>
                                <td className="border px-2 py-1">
                                  {[
                                    farmer.village,
                                    farmer.union_name,
                                    farmer.upazila,
                                    farmer.district,
                                  ]
                                    .filter(Boolean)
                                    .join(", ")}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Category
                                </td>
                                <td className="border px-2 py-1 font-bold">
                                  {farmer.farm_category}
                                </td>
                              </tr>
                              <tr>
                                <td className="border px-2 py-1 font-semibold bg-slate-50">
                                  Quantities
                                </td>
                                <td className="border px-2 py-1 font-semibold text-blue-700">
                                  {breakdown.length === 0
                                    ? "-"
                                    : breakdown
                                        .map((b) => `${b.label}: ${b.qty}`)
                                        .join(", ")}
                                </td>
                              </tr>
                              {farmer.feed_type && (
                                <tr>
                                  <td className="border px-2 py-1 font-semibold bg-slate-50">
                                    Feed
                                  </td>
                                  <td className="border px-2 py-1">
                                    {farmer.feed_type}
                                  </td>
                                </tr>
                              )}
                            </tbody>
                          </table>
                        </div>
                      </Popup>
                    </Marker>
                  );
                })}
            </MapContainer>

            {/* Floating Locate Me Button */}
            <button
              onClick={() => {
                navigator.geolocation.getCurrentPosition((position) => {
                  setSelectedFarmerPosition(null);
                  setMyPosition([
                    position.coords.latitude,
                    position.coords.longitude,
                  ]);
                  setMyLocationMode(true);
                });
              }}
              className="absolute top-3 left-1/2 -translate-x-1/2 z-[500] flex items-center gap-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-full shadow-lg px-3.5 py-1.5 text-xs font-semibold transition"
            >
              <MapPin size={15} /> Locate Me
            </button>

            {/* Floating Top-Right Close Map Icon */}
            <button
              type="button"
              onClick={() => setShowFarmerMap(false)}
              title="Close Map"
              className="absolute top-3 right-3 z-[500] flex items-center gap-1 bg-red-600 hover:bg-red-700 text-white rounded-full shadow-lg px-3 py-1.5 text-xs font-bold transition"
            >
              <X size={14} /> Close
            </button>
          </div>
        </div>
      )}

      {/* 4. AGENT-WISE FARMER TABLE */}
      {agentGroupedFarmers.length === 0 ? (
        <div className="bg-white border border-slate-300 rounded-lg p-8 text-center text-xs text-slate-500">
          No farmers found for the selected filter.
        </div>
      ) : (
        <div className="space-y-4">
          {agentGroupedFarmers.map((group: any) => (
            <div
              key={group.agentId}
              className="bg-white border border-slate-300 rounded-lg overflow-hidden shadow-sm"
            >
              <div className="bg-slate-100 border-b border-slate-300 px-3 py-2 flex flex-wrap items-center justify-between gap-2">
                <div className="text-xs sm:text-sm font-bold text-slate-800">
                  Agent: {group.agentName}{" "}
                  <span className="font-normal text-slate-600">
                    ({group.companyName})
                  </span>
                </div>
                <span className="text-xs font-semibold bg-white border border-slate-300 px-2 py-0.5 rounded text-slate-700">
                  Farmers: {group.farmers.length}
                </span>
              </div>

              <div className="overflow-x-auto">
                <table className="w-full border-collapse text-xs">
                  <thead className="bg-slate-50">
                    <tr>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        SL
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Farmer
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Mobile
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Village
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Union
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Upazila
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Category
                      </th>
                      <th className="border border-slate-300 px-2 py-1.5 text-left">
                        Farm Details
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
                    {group.farmers.map((farmer: any, idx: number) => {
                      const breakdown = getStockBreakdownList(farmer);
                      return (
                        <tr key={farmer.id} className="hover:bg-slate-50">
                          <td className="border border-slate-300 px-2 py-1">
                            {idx + 1}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 font-semibold text-slate-800">
                            {farmer.farmer_name}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {farmer.mobile || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {farmer.village || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {farmer.union_name || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {farmer.upazila || "-"}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 font-medium">
                            {farmer.farm_category}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {breakdown.length === 0
                              ? "-"
                              : breakdown
                                  .map((b) => `${b.label}: ${b.qty}`)
                                  .join(", ")}
                          </td>
                          <td className="border border-slate-300 px-2 py-1">
                            {farmer.status}
                          </td>
                          <td className="border border-slate-300 px-2 py-1 text-center whitespace-nowrap space-x-1">
                            <button
                              onClick={() => setSelectedFarmer(farmer)}
                              className="bg-blue-600 hover:bg-blue-700 text-white px-2 py-0.5 rounded text-[11px]"
                            >
                              View
                            </button>
                            <button
                              onClick={() => handleViewFarmerOnMap(farmer)}
                              className="bg-teal-600 hover:bg-teal-700 text-white px-2 py-0.5 rounded text-[11px]"
                              title="View on Map"
                            >
                              View on Map
                            </button>
                            <button
                              onClick={() => openEditForm(farmer)}
                              className="bg-yellow-500 hover:bg-yellow-600 text-white px-2 py-0.5 rounded text-[11px]"
                            >
                              Edit
                            </button>
                            <button
                              onClick={() => handleDelete(farmer.id)}
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

      {/* 5. VIEW FARMER MODAL (z-[2000] so it always renders ABOVE Leaflet map) */}
      {selectedFarmer && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[2000] p-2 sm:p-4">
          <div className="bg-white rounded-xl w-full max-w-xl max-h-[90vh] flex flex-col overflow-hidden shadow-2xl border border-slate-300">
            <div className="flex items-center justify-between px-4 py-2.5 border-b border-slate-300 bg-slate-100">
              <h2 className="text-sm sm:text-base font-bold text-slate-800">
                Farmer: {selectedFarmer.farmer_name}
              </h2>
              <button
                onClick={() => setSelectedFarmer(null)}
                className="text-slate-500 hover:text-slate-800"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-4 space-y-3 overflow-y-auto">
              <div className="bg-emerald-50 border border-emerald-300 rounded-lg p-3 space-y-2">
                <div className="flex items-center justify-between text-xs">
                  <span className="font-bold text-emerald-900">
                    📝 Farmer Tracking
                  </span>
                  <span className="text-emerald-700">
                    User: <b>{currentUserName}</b>
                  </span>
                </div>

                <div className="max-h-36 overflow-y-auto space-y-1.5">
                  {loadingNotes ? (
                    <div className="text-xs text-slate-500">Loading...</div>
                  ) : farmerNotes.length === 0 ? (
                    <div className="text-xs text-slate-400 bg-white border border-emerald-200 rounded p-2 text-center">
                      No tracking notes yet.
                    </div>
                  ) : (
                    farmerNotes.map((item) => (
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
                  onSubmit={(e) => handleAddFarmerNote(e, selectedFarmer.id)}
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
                      {selectedFarmer.companies?.company_name || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Agent
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.agents?.agent_name || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Farmer Name
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-bold">
                      {selectedFarmer.farmer_name}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Mobile
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.mobile || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Address
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {[
                        selectedFarmer.village,
                        selectedFarmer.union_name,
                        selectedFarmer.upazila,
                        selectedFarmer.district,
                      ]
                        .filter(Boolean)
                        .join(", ")}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Farm Category
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-semibold">
                      {selectedFarmer.farm_category}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Stock / Capacity
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5 font-semibold text-blue-700">
                      {getStockBreakdownList(selectedFarmer).length === 0
                        ? "-"
                        : getStockBreakdownList(selectedFarmer)
                            .map((b) => `${b.label}: ${b.qty}`)
                            .join(", ")}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Feed Type
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.feed_type || "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Status
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.status || "Active"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Coordinates
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.latitude && selectedFarmer.longitude
                        ? `${selectedFarmer.latitude}, ${selectedFarmer.longitude}`
                        : "-"}
                    </td>
                  </tr>
                  <tr>
                    <td className="border border-slate-300 bg-slate-50 px-3 py-1.5 font-semibold">
                      Remarks
                    </td>
                    <td className="border border-slate-300 px-3 py-1.5">
                      {selectedFarmer.remarks || "-"}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* 6. ADD / EDIT FARMER MODAL (Company -> District -> Upazila -> Agent -> Rest) */}
      {showForm && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[2000] p-2 sm:p-4">
          <div className="bg-white rounded-xl p-4 sm:p-5 w-full max-w-2xl max-h-[92vh] overflow-y-auto shadow-2xl">
            <div className="flex justify-between items-center pb-2.5 mb-3 border-b">
              <h2 className="text-base font-bold text-slate-800">
                {editingId ? "Edit Farmer" : "Add Farmer"}
              </h2>
              <button onClick={() => setShowForm(false)}>
                <X size={18} />
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
              {/* 1. Company (Default Nourish) */}
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Company *
                </label>
                <select
                  value={form.company_id}
                  onChange={(e) => {
                    const companyId = e.target.value;
                    setForm({ ...form, company_id: companyId, agent_id: "" });
                  }}
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

              {/* 2. District (Assigned Districts Chronologically) */}
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
                      agent_id: "",
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

              {/* 3. Upazila (Under Selected District) */}
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
                      agent_id: "",
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

              {/* 4. Agent (Filtered by Selected Company & Upazila) */}
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Agent *
                </label>
                <select
                  value={form.agent_id}
                  onChange={(e) =>
                    setForm({ ...form, agent_id: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Agent</option>
                  {formAvailableAgents.map((a) => (
                    <option key={a.id} value={a.id}>
                      {a.agent_name}
                    </option>
                  ))}
                </select>
              </div>

              {/* 5. Remaining Fields (District & Upazila not repeated) */}
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Farmer Name *
                </label>
                <input
                  placeholder="Farmer Name"
                  value={form.farmer_name}
                  onChange={(e) =>
                    setForm({ ...form, farmer_name: e.target.value })
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
                  Union
                </label>
                <select
                  value={form.union_name}
                  onChange={(e) =>
                    setForm({ ...form, union_name: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
                >
                  <option value="">Select Union</option>
                  {unions.map((u) => (
                    <option key={u.id} value={u.name}>
                      {u.name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Village
                </label>
                <input
                  placeholder="Village"
                  value={form.village}
                  onChange={(e) =>
                    setForm({ ...form, village: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>
            </div>

            <div className="mb-3 flex flex-wrap items-center gap-2 bg-slate-50 border border-slate-200 rounded p-2.5">
              <button
                type="button"
                onClick={() => {
                  setPickerCenterOverride(null);
                  setShowMap(true);
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

            <div className="mb-3">
              <label className="block text-xs font-bold text-slate-600 mb-1">
                Farm Category *
              </label>
              <select
                value={form.farm_category || ""}
                onChange={(e) =>
                  setForm({ ...form, farm_category: e.target.value })
                }
                className="w-full rounded border border-slate-300 bg-white px-2.5 py-1.5 text-xs"
              >
                <option value="">Select Category</option>
                <option value="Poultry">🐔 Poultry</option>
                <option value="Cattle">🐄 Cattle</option>
                <option value="Fish">🐟 Fish</option>
              </select>
            </div>

            {form.farm_category === "Poultry" && (
              <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 mb-3 bg-amber-50 border border-amber-200 rounded p-2.5">
                {POULTRY_FIELDS.map((item) => (
                  <div key={item}>
                    <label className="block text-[11px] font-semibold capitalize">
                      {item.replace("_", " ")}
                    </label>
                    <input
                      type="number"
                      min="0"
                      value={(form as any)[item] || ""}
                      onChange={(e) =>
                        setForm({ ...form, [item]: e.target.value })
                      }
                      className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                    />
                  </div>
                ))}
              </div>
            )}

            {form.farm_category === "Cattle" && (
              <div className="grid grid-cols-3 gap-2 mb-3 bg-emerald-50 border border-emerald-200 rounded p-2.5">
                {CATTLE_FIELDS.map((item) => (
                  <div key={item}>
                    <label className="block text-[11px] font-semibold capitalize">
                      {item.replace("_", " ")}
                    </label>
                    <input
                      type="number"
                      min="0"
                      value={(form as any)[item] || ""}
                      onChange={(e) =>
                        setForm({ ...form, [item]: e.target.value })
                      }
                      className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                    />
                  </div>
                ))}
              </div>
            )}

            {form.farm_category === "Fish" && (
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 mb-3 bg-sky-50 border border-sky-200 rounded p-2.5">
                {FISH_FIELDS.map((item) => (
                  <div key={item}>
                    <label className="block text-[11px] font-semibold capitalize">
                      {item}
                    </label>
                    <input
                      type="number"
                      min="0"
                      value={(form as any)[item] || ""}
                      onChange={(e) =>
                        setForm({ ...form, [item]: e.target.value })
                      }
                      className="w-full rounded border border-slate-300 bg-white px-2 py-1 text-xs"
                    />
                  </div>
                ))}
              </div>
            )}

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-4">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Feed Type
                </label>
                <input
                  placeholder="Feed Type"
                  value={form.feed_type}
                  onChange={(e) =>
                    setForm({ ...form, feed_type: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Remarks
                </label>
                <input
                  placeholder="Remarks"
                  value={form.remarks}
                  onChange={(e) =>
                    setForm({ ...form, remarks: e.target.value })
                  }
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs"
                />
              </div>
            </div>

            <div className="flex justify-end gap-2 pt-2 border-t">
              <button
                onClick={() => setShowForm(false)}
                className="px-4 py-1.5 rounded border border-slate-300 text-xs font-semibold"
              >
                Cancel
              </button>
              <button
                onClick={handleSave}
                className="px-4 py-1.5 rounded bg-green-600 text-white text-xs font-semibold"
              >
                {editingId ? "Update Farmer" : "Save Farmer"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 7. LOCATION PICKER MODAL */}
      {showMap && (
        <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-[2100] p-2 sm:p-4">
          <div className="bg-white p-3 sm:p-4 rounded-xl w-full max-w-4xl shadow-2xl space-y-2.5">
            <div className="flex flex-wrap items-center justify-between gap-2">
              <h2 className="font-bold text-sm sm:text-base">
                Select Farm Location
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
                    onClick={() => setShowMap(false)}
                    className="inline-flex items-center gap-1 bg-green-600 hover:bg-green-700 text-white px-3 py-1.5 rounded text-xs font-semibold"
                  >
                    <CheckCircle2 size={13} /> Use Selected Location
                  </button>
                )}

                <button onClick={() => setShowMap(false)}>
                  <X size={18} />
                </button>
              </div>
            </div>

            <LocationPicker
              key={
                pickerCenterOverride
                  ? `${pickerCenterOverride[0]}_${pickerCenterOverride[1]}`
                  : "default_picker"
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
                setShowMap(false);
              }}
              markers={
                <>
                  {farmers
                    .filter((f) => f.latitude && f.longitude)
                    .map((farmer) => (
                      <Marker
                        key={farmer.id}
                        position={[
                          Number(farmer.latitude),
                          Number(farmer.longitude),
                        ]}
                        icon={buildCustomFarmMarker(farmer)}
                      />
                    ))}
                </>
              }
            />
          </div>
        </div>
      )}

      {/* 8. FARMER SUMMARY MODAL */}
      {showSummary && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-[2000] p-2 sm:p-4">
          <div className="bg-white rounded-xl p-4 sm:p-6 w-full max-w-6xl max-h-[90vh] overflow-auto">
            <div className="flex justify-between items-center pb-3 mb-4 border-b">
              <div>
                <h2 className="text-base sm:text-lg font-bold">
                  Farmer Summary — {selectedCompanyName}
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

            <div className="mb-4">
              <select
                defaultValue=""
                onChange={(e) => {
                  const val = e.target.value;
                  if (val === "company") exportFarmerSummaryExcel();
                  if (val === "poultry") exportPoultrySummaryExcel();
                  if (val === "cattle") exportCattleSummaryExcel();
                  if (val === "fish") exportFishSummaryExcel();
                  e.target.value = "";
                }}
                className="bg-slate-800 text-white px-3.5 py-1.5 rounded text-xs font-semibold cursor-pointer"
              >
                <option value="" hidden>
                  📊 Export Summary
                </option>
                <option value="company">Company Farmer Summary</option>
                <option value="poultry">🐔 Poultry Summary</option>
                <option value="cattle">🐄 Cattle Summary</option>
                <option value="fish">🐟 Fish Summary</option>
              </select>
            </div>

            <div className="overflow-x-auto mb-6">
              <table className="w-full border-collapse text-xs">
                <thead className="bg-slate-100">
                  <tr>
                    <th className="border px-3 py-1.5 text-left">Company</th>
                    <th className="border px-3 py-1.5 text-center">
                      Total Farmers
                    </th>
                    <th className="border px-3 py-1.5 text-center">
                      🐔 Poultry Farmers
                    </th>
                    <th className="border px-3 py-1.5 text-center">
                      🐄 Cattle Farmers
                    </th>
                    <th className="border px-3 py-1.5 text-center">
                      🐟 Fish Farmers
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {companySummary.map((row: any) => (
                    <tr key={row.company}>
                      <td className="border px-3 py-1.5 font-semibold">
                        {row.company}
                      </td>
                      <td className="border px-3 py-1.5 text-center font-bold">
                        {row.totalFarmers}
                      </td>
                      <td className="border px-3 py-1.5 text-center text-amber-700 font-semibold">
                        {row.poultryFarmers}
                      </td>
                      <td className="border px-3 py-1.5 text-center text-emerald-700 font-semibold">
                        {row.cattleFarmers}
                      </td>
                      <td className="border px-3 py-1.5 text-center text-sky-700 font-semibold">
                        {row.fishFarmers}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <h3 className="text-sm font-bold mb-2">🐔 Poultry Summary</h3>
            <div className="overflow-x-auto mb-6">
              <table className="w-full border-collapse text-xs">
                <thead className="bg-amber-100">
                  <tr>
                    <th className="border px-2 py-1.5 text-left">Company</th>
                    <th className="border px-2 py-1.5">Broiler Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Broiler
                    </th>
                    <th className="border px-2 py-1.5">Layer Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Layer
                    </th>
                    <th className="border px-2 py-1.5">Sonali Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Sonali
                    </th>
                    <th className="border px-2 py-1.5">Duck Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Duck
                    </th>
                    <th className="border px-2 py-1.5">Quail Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Quail
                    </th>
                    <th className="border px-2 py-1.5">Deshi Farms</th>
                    <th className="border px-2 py-1.5 bg-amber-50">
                      Total Deshi
                    </th>
                    <th className="border px-2 py-1.5 bg-amber-200 font-bold">
                      Total Poultry
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {companySummary.map((row: any) => (
                    <tr key={row.company}>
                      <td className="border px-2 py-1 font-semibold">
                        {row.company}
                      </td>
                      <td className="border text-center">
                        {row.broilerFarmers}
                      </td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.broiler}
                      </td>
                      <td className="border text-center">{row.layerFarmers}</td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.layer}
                      </td>
                      <td className="border text-center">
                        {row.sonaliFarmers}
                      </td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.sonali}
                      </td>
                      <td className="border text-center">{row.duckFarmers}</td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.duck}
                      </td>
                      <td className="border text-center">{row.quailFarmers}</td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.quail}
                      </td>
                      <td className="border text-center">{row.deshiFarmers}</td>
                      <td className="border text-center bg-amber-50 font-semibold">
                        {row.deshiChicken}
                      </td>
                      <td className="border text-center bg-amber-100 font-bold">
                        {row.totalPoultry}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <h3 className="text-sm font-bold mb-2">🐄 Cattle Summary</h3>
            <div className="overflow-x-auto">
              <table className="w-full border-collapse text-xs">
                <thead className="bg-emerald-100">
                  <tr>
                    <th className="border px-2 py-1.5 text-left">Company</th>
                    <th className="border px-2 py-1.5">Dairy Farms</th>
                    <th className="border px-2 py-1.5 bg-emerald-50">
                      Total Dairy Cow
                    </th>
                    <th className="border px-2 py-1.5">Bull Farms</th>
                    <th className="border px-2 py-1.5 bg-emerald-50">
                      Total Bull
                    </th>
                    <th className="border px-2 py-1.5">Calf/Heifer Farms</th>
                    <th className="border px-2 py-1.5 bg-emerald-50">
                      Total Calf/Heifer
                    </th>
                    <th className="border px-2 py-1.5 bg-emerald-200 font-bold">
                      Total Cattle
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {companySummary.map((row: any) => (
                    <tr key={row.company}>
                      <td className="border px-2 py-1 font-semibold">
                        {row.company}
                      </td>
                      <td className="border text-center">
                        {row.dairyCowFarmers}
                      </td>
                      <td className="border text-center bg-emerald-50 font-semibold">
                        {row.dairyCow}
                      </td>
                      <td className="border text-center">{row.bullFarmers}</td>
                      <td className="border text-center bg-emerald-50 font-semibold">
                        {row.bull}
                      </td>
                      <td className="border text-center">
                        {row.calfHeiferFarmers}
                      </td>
                      <td className="border text-center bg-emerald-50 font-semibold">
                        {row.calfHeifer}
                      </td>
                      <td className="border text-center bg-emerald-100 font-bold">
                        {row.totalCattle}
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