import React, { useEffect, useRef, useState } from "react";
import { supabase } from "../../lib/supabase";
import { useAuth } from "../../context/AuthContext";
import SearchableSelect from "../../components/common/SearchableSelect";
import { isManagerOrAdmin } from "../../utils/accessControl";
import { FaMap, FaMapMarkerAlt, FaLayerGroup, FaStore, FaUndo } from "react-icons/fa";

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

// In-memory page cache (same pattern as FarmersPage / DocTrackingPage):
// returning to this page shows the last loaded data instantly with no
// reload. The cache lives only for the session; a full refresh loads
// fresh. Access inputs are cached too (userId, assignedDistricts,
// assignedUpazilas) — the officer scoping depends on them, so a cache
// hit must not leave them empty. Filters are NOT cached — they start
// at "all" on every mount.
let cachedPageState: {
  initialized: boolean;
  userId: string;
  assignedDistricts: string[];
  assignedUpazilas: string[];
  districts: any[];
  upazilas: any[];
  unions: any[];
  bazars: any[];
} | null = null;

type TabKey = "district" | "upazila" | "union" | "bazar";

export default function AreaPage() {
  const { role } = useAuth();
  const isFullAccess = role === "super_admin" || role === "manager";
  const canDeleteBazar = isManagerOrAdmin(role);

  const [, setCurrentUserId] = useState(
    cachedPageState?.userId || ""
  );
  const [assignedDistricts, setAssignedDistricts] = useState<string[]>(
    cachedPageState?.assignedDistricts || []
  );
  const [assignedUpazilas, setAssignedUpazilas] = useState<string[]>(
    cachedPageState?.assignedUpazilas || []
  );

  const [districts, setDistricts] = useState<any[]>(
    cachedPageState?.districts || []
  );
  const [upazilas, setUpazilas] = useState<any[]>(
    cachedPageState?.upazilas || []
  );
  const [unions, setUnions] = useState<any[]>(
    cachedPageState?.unions || []
  );
  const [bazars, setBazars] = useState<any[]>(
    cachedPageState?.bazars || []
  );

  const [loading, setLoading] = useState(!cachedPageState?.initialized);

  // Filters are NOT cached — they start at "all" on every mount.
  const [activeTab, setActiveTab] = useState<TabKey>("district");
  const [districtFilter, setDistrictFilter] = useState("");
  const [upazilaFilter, setUpazilaFilter] = useState("");
  const [unionFilter, setUnionFilter] = useState("");

  // Tracks whether the default district has been applied, so a deliberate
  // "All Districts" clear is never re-applied behind the user's back.
  const defaultDistrictAppliedRef = useRef(false);

  // Bazar Add/Edit modal state.
  const [showModal, setShowModal] = useState(false);
  const [editingBazarId, setEditingBazarId] = useState<number | null>(null);
  const [bazarName, setBazarName] = useState("");
  const [modalDistrict, setModalDistrict] = useState("");
  const [modalUpazila, setModalUpazila] = useState("");
  const [modalUnion, setModalUnion] = useState("");
  const [savingBazar, setSavingBazar] = useState(false);

  // ---------- Visible (scoped) rows ----------
  // Officer sees only their assigned areas; manager / super admin see
  // everything. An EMPTY assigned list means ALL (the app's established
  // convention, same as FarmersPage). Unions derive from the visible
  // upazilas; bazars derive from the visible unions.
  const visibleDistricts = isFullAccess
    ? districts
    : districts.filter(
        (d: any) =>
          assignedDistricts.length === 0 ||
          assignedDistricts.includes(d.name)
      );

  const visibleDistrictIds = new Set(
    visibleDistricts.map((d: any) => String(d.id))
  );

  const visibleUpazilas = upazilas.filter((u: any) => {
    if (!visibleDistrictIds.has(String(u.district_id))) return false;
    if (isFullAccess) return true;
    return (
      assignedUpazilas.length === 0 || assignedUpazilas.includes(u.name)
    );
  });

  const visibleUpazilaIds = new Set(
    visibleUpazilas.map((u: any) => String(u.id))
  );

  const visibleUnions = unions.filter((un: any) =>
    visibleUpazilaIds.has(String(un.upazila_id))
  );

  const visibleUnionIds = new Set(
    visibleUnions.map((un: any) => String(un.id))
  );

  const visibleBazars = bazars.filter((b: any) =>
    visibleUnionIds.has(String(b.union_id))
  );

  // ---------- Lookups ----------
  const districtById: Record<string, any> = {};
  districts.forEach((d: any) => {
    districtById[String(d.id)] = d;
  });
  const upazilaById: Record<string, any> = {};
  upazilas.forEach((u: any) => {
    upazilaById[String(u.id)] = u;
  });
  const unionById: Record<string, any> = {};
  unions.forEach((un: any) => {
    unionById[String(un.id)] = un;
  });

  // ---------- Filtered rows per tab ----------
  const selectedFilterDistrict = districtFilter
    ? districts.find((d: any) => d.name === districtFilter)
    : null;
  const selectedFilterUpazila = upazilaFilter
    ? visibleUpazilas.find((u: any) => u.name === upazilaFilter)
    : null;

  const filterUpazilas = selectedFilterDistrict
    ? visibleUpazilas.filter(
        (u: any) => String(u.district_id) === String(selectedFilterDistrict.id)
      )
    : visibleUpazilas;

  const filterUnions = selectedFilterUpazila
    ? visibleUnions.filter(
        (un: any) => String(un.upazila_id) === String(selectedFilterUpazila.id)
      )
    : selectedFilterDistrict
      ? visibleUnions.filter((un: any) => {
          const upa = upazilaById[String(un.upazila_id)];
          return (
            upa && String(upa.district_id) === String(selectedFilterDistrict.id)
          );
        })
      : visibleUnions;

  const districtRows = visibleDistricts;

  const upazilaRows = selectedFilterDistrict
    ? visibleUpazilas.filter(
        (u: any) => String(u.district_id) === String(selectedFilterDistrict.id)
      )
    : visibleUpazilas;

  const unionRows = filterUnions;

  const bazarRows = visibleBazars.filter((b: any) => {
    if (unionFilter) return String(b.union_id) === String(unionFilter);
    if (selectedFilterUpazila)
      return String(
        unionById[String(b.union_id)]?.upazila_id || ""
      ) === String(selectedFilterUpazila.id);
    if (selectedFilterDistrict) {
      const upa = upazilaById[String(unionById[String(b.union_id)]?.upazila_id || "")];
      return upa && String(upa.district_id) === String(selectedFilterDistrict.id);
    }
    return true;
  });

  // Default district: once the lists are loaded, pre-select the first
  // visible district (the lists arrive ordered by name, so this is the
  // alphabetically first one — for a single-district officer, their
  // district). The ref keeps a deliberate "All Districts" clear from
  // being re-applied. Placed here — after visibleDistricts is declared —
  // so the dep array never reads it before declaration.
  useEffect(() => {
    if (loading) return;
    if (defaultDistrictAppliedRef.current) return;
    if (districtFilter !== "") return;
    if (visibleDistricts.length === 0) return;
    defaultDistrictAppliedRef.current = true;
    setDistrictFilter(visibleDistricts[0].name);
  }, [loading, districtFilter, visibleDistricts]);

  // Reset filters: back to the initial state — default district selected,
  // upazila and union cleared.
  function resetFilters() {
    setDistrictFilter(
      visibleDistricts.length > 0 ? visibleDistricts[0].name : ""
    );
    setUpazilaFilter("");
    setUnionFilter("");
    defaultDistrictAppliedRef.current = true;
  }

  // ---------- Modal cascading options (scoped the same way) ----------
  const modalDistrictObj = modalDistrict
    ? visibleDistricts.find((d: any) => d.name === modalDistrict)
    : null;
  const modalUpazilas = modalDistrictObj
    ? visibleUpazilas.filter(
        (u: any) => String(u.district_id) === String(modalDistrictObj.id)
      )
    : visibleUpazilas;
  const modalUpazilaObj = modalUpazila
    ? modalUpazilas.find((u: any) => u.name === modalUpazila)
    : null;
  const modalUnions = modalUpazilaObj
    ? visibleUnions.filter(
        (un: any) => String(un.upazila_id) === String(modalUpazilaObj.id)
      )
    : [];

  async function loadData() {
    const {
      data: { user },
    } = await supabase.auth.getUser();

    let uId = user?.id || "";
    let uDistricts: string[] = [];
    let uUpazilas: string[] = [];

    if (user) {
      uId = user.id;
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
        uDistricts = parseList(profile.districts);
        uUpazilas = parseList(profile.upazilas);
      }
    }

    setAssignedDistricts(uDistricts);
    setAssignedUpazilas(uUpazilas);

    const [
      { data: districtData },
      { data: upazilaData },
      { data: unionData },
      { data: bazarData },
    ] = await Promise.all([
      supabase.from("districts").select("*").order("name"),
      supabase.from("upazilas").select("*").order("name"),
      supabase.from("unions").select("*").order("name"),
      supabase.from("bazars").select("*").order("name"),
    ]);

    const loadedDistricts = districtData || [];
    const loadedUpazilas = upazilaData || [];
    const loadedUnions = unionData || [];
    const loadedBazars = bazarData || [];

    setDistricts(loadedDistricts);
    setUpazilas(loadedUpazilas);
    setUnions(loadedUnions);
    setBazars(loadedBazars);

    cachedPageState = {
      initialized: true,
      userId: uId,
      assignedDistricts: uDistricts,
      assignedUpazilas: uUpazilas,
      districts: loadedDistricts,
      upazilas: loadedUpazilas,
      unions: loadedUnions,
      bazars: loadedBazars,
    };
  }

  // Mount: validate the session user against the cache — if the user
  // changed (e.g. account switched in the same tab), the cached lists
  // belong to someone else, so drop them and load fresh. A stale cache
  // must never serve another user's data (2026-11-20 lesson).
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const {
          data: { user },
        } = await supabase.auth.getUser();
        if (cancelled) return;
        if (
          cachedPageState?.initialized &&
          user &&
          cachedPageState.userId === user.id
        ) {
          setLoading(false);
          return;
        }
        cachedPageState = null;
        setDistricts([]);
        setUpazilas([]);
        setUnions([]);
        setBazars([]);
        setAssignedDistricts([]);
        setAssignedUpazilas([]);
        await loadData();
      } catch {
        // Load failure leaves the empty state; the user can retry by
        // leaving and returning to the page.
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function openAddBazar() {
    setEditingBazarId(null);
    setBazarName("");
    // Prefill the cascade from the current table filters when set.
    setModalDistrict(districtFilter || "");
    setModalUpazila(upazilaFilter || "");
    setModalUnion(unionFilter || "");
    setShowModal(true);
  }

  function openEditBazar(bazar: any) {
    const union = unionById[String(bazar.union_id)];
    const upazila = union ? upazilaById[String(union.upazila_id)] : null;
    const district = upazila
      ? districtById[String(upazila.district_id)]
      : null;
    setEditingBazarId(Number(bazar.id));
    setBazarName(bazar.name || "");
    setModalDistrict(district?.name || "");
    setModalUpazila(upazila?.name || "");
    setModalUnion(String(bazar.union_id));
    setShowModal(true);
  }

  async function handleSaveBazar() {
    if (savingBazar) return;
    const name = bazarName.trim();
    if (!name) {
      alert("Bazar name is required.");
      return;
    }
    if (!modalUnion) {
      alert("District, Upazila and Union are required.");
      return;
    }
    setSavingBazar(true);
    try {
      if (editingBazarId) {
        const { data, error } = await supabase
          .from("bazars")
          .update({ name, union_id: Number(modalUnion) })
          .eq("id", editingBazarId)
          .select()
          .single();
        if (error) throw error;
        const updated = bazars.map((b: any) =>
          Number(b.id) === editingBazarId ? data : b
        );
        setBazars(updated);
        if (cachedPageState) cachedPageState.bazars = updated;
      } else {
        const { data, error } = await supabase
          .from("bazars")
          .insert([{ name, union_id: Number(modalUnion) }])
          .select()
          .single();
        if (error) throw error;
        const updated = [...bazars, data].sort((a: any, b: any) =>
          String(a.name || "").localeCompare(String(b.name || ""))
        );
        setBazars(updated);
        if (cachedPageState) cachedPageState.bazars = updated;
      }
      setShowModal(false);
    } catch (err: any) {
      alert(err?.message || "Failed to save bazar.");
    } finally {
      setSavingBazar(false);
    }
  }

  async function handleDeleteBazar(bazar: any) {
    // Backstop guard: the Delete button is hidden for officers too.
    if (!canDeleteBazar) {
      alert("Only managers and super admins can delete a bazar.");
      return;
    }
    if (!window.confirm(`Delete bazar "${bazar.name}"?`)) return;
    const { error } = await supabase
      .from("bazars")
      .delete()
      .eq("id", bazar.id);
    if (error) {
      alert(error.message || "Failed to delete bazar.");
      return;
    }
    const updated = bazars.filter((b: any) => Number(b.id) !== Number(bazar.id));
    setBazars(updated);
    if (cachedPageState) cachedPageState.bazars = updated;
  }

  const tabs: { key: TabKey; label: string; count: number; icon: React.ReactNode }[] = [
    {
      key: "district",
      label: "District",
      count: visibleDistricts.length,
      icon: <FaMap className="text-emerald-600" size={18} />,
    },
    {
      key: "upazila",
      label: "Upazila",
      count: visibleUpazilas.length,
      icon: <FaMapMarkerAlt className="text-sky-600" size={18} />,
    },
    {
      key: "union",
      label: "Union",
      count: visibleUnions.length,
      icon: <FaLayerGroup className="text-amber-600" size={18} />,
    },
    {
      key: "bazar",
      label: "Bazar",
      count: visibleBazars.length,
      icon: <FaStore className="text-violet-600" size={18} />,
    },
  ];

  function renderTable(
    headers: string[],
    rows: React.ReactNode,
    emptyText: string
  ) {
    return (
      <div className="overflow-x-auto">
        <table className="w-full border-collapse text-xs">
          <thead className="bg-slate-200">
            <tr>
              {headers.map((h, i) => (
                <th
                  key={h}
                  className={`border border-slate-300 px-2 py-1.5 ${
                    i === headers.length - 1 && h === "Action"
                      ? "text-center"
                      : "text-left"
                  }`}
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows}
            {((activeTab === "district" && districtRows.length === 0) ||
              (activeTab === "upazila" && upazilaRows.length === 0) ||
              (activeTab === "union" && unionRows.length === 0) ||
              (activeTab === "bazar" && bazarRows.length === 0)) && (
              <tr>
                <td
                  colSpan={headers.length}
                  className="border border-slate-300 px-2 py-3 text-center text-slate-500"
                >
                  {emptyText}
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    );
  }

  return (
    <div className="p-3 sm:p-5 max-w-7xl mx-auto space-y-3">
      <div className="bg-white border border-slate-300 rounded-xl px-3.5 py-2.5 shadow-sm">
        <h1 className="text-sm sm:text-base font-bold text-slate-800">
          Area
        </h1>
        <p className="text-xs text-slate-500">
          Districts, upazilas, unions and bazars in your area.
        </p>
      </div>

      {/* 1. TAB CARDS */}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        {tabs.map((tab) => (
          <button
            key={tab.key}
            onClick={() => setActiveTab(tab.key)}
            className={`flex items-center gap-2.5 rounded-xl border px-3 py-2.5 text-left shadow-sm transition ${
              activeTab === tab.key
                ? "bg-emerald-50 border-emerald-600 ring-2 ring-emerald-600"
                : "bg-white border-slate-300 hover:bg-slate-50"
            }`}
          >
            {tab.icon}
            <span className="min-w-0">
              <span className="block text-xs font-bold text-slate-800">
                {tab.label}
              </span>
              <span className="block text-xs text-slate-500">
                {tab.count} {tab.count === 1 ? "item" : "items"}
              </span>
            </span>
          </button>
        ))}
      </div>

      {/* 2. LIST CARD */}
      <div className="bg-white border border-slate-300 rounded-xl shadow-sm overflow-hidden">
        <div className="bg-slate-100 border-b border-slate-300 px-3 py-2 flex flex-wrap items-center justify-between gap-2">
          <div className="text-xs sm:text-sm font-bold text-slate-800">
            {activeTab === "district" && `Districts (${districtRows.length})`}
            {activeTab === "upazila" && `Upazilas (${upazilaRows.length})`}
            {activeTab === "union" && `Unions (${unionRows.length})`}
            {activeTab === "bazar" && `Bazars (${bazarRows.length})`}
          </div>
          {activeTab === "bazar" && (
            <button
              onClick={openAddBazar}
              className="inline-flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-lg text-xs font-semibold shadow-sm transition"
            >
              + Add Bazar
            </button>
          )}
        </div>

        {/* Filters */}
        {activeTab !== "district" && (
          <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 p-2.5 border-b border-slate-200">
            <SearchableSelect
              searchable={false}
              value={districtFilter}
              onChange={(v) => {
                setDistrictFilter(v);
                setUpazilaFilter("");
                setUnionFilter("");
              }}
              placeholder="All Districts"
              options={visibleDistricts.map((d: any) => ({
                value: d.name,
                label: d.name,
              }))}
            />
            {(activeTab === "upazila" ||
              activeTab === "union" ||
              activeTab === "bazar") && (
              <SearchableSelect
                value={upazilaFilter}
                onChange={(v) => {
                  setUpazilaFilter(v);
                  setUnionFilter("");
                }}
                placeholder="All Upazilas"
                options={filterUpazilas.map((u: any) => ({
                  value: u.name,
                  label: u.name,
                }))}
              />
            )}
            {(activeTab === "union" || activeTab === "bazar") && (
              <SearchableSelect
                value={unionFilter}
                onChange={setUnionFilter}
                placeholder="All Unions"
                options={filterUnions.map((un: any) => ({
                  value: String(un.id),
                  label: un.name,
                }))}
              />
            )}
            <button
              type="button"
              onClick={resetFilters}
              title="Reset filters"
              className="inline-flex items-center justify-center gap-1.5 border border-slate-300 bg-white hover:bg-slate-100 text-slate-700 px-3 py-1.5 rounded-lg text-xs font-semibold transition"
            >
              <FaUndo size={12} />
              Reset
            </button>
          </div>
        )}

        {loading ? (
          <div className="p-8 text-center text-xs text-slate-500">
            Loading...
          </div>
        ) : (
          <>
            {activeTab === "district" &&
              renderTable(
                ["SL", "District"],
                districtRows.map((d: any, idx: number) => (
                  <tr
                    key={d.id}
                    className={`${
                      idx % 2 === 0 ? "bg-white" : "bg-slate-100"
                    } hover:bg-emerald-100`}
                  >
                    <td className="border border-slate-300 px-2 py-1.5">
                      {idx + 1}
                    </td>
                    <td className="border border-slate-300 px-2 py-1.5 font-semibold">
                      {d.name}
                    </td>
                  </tr>
                )),
                "No districts found."
              )}

            {activeTab === "upazila" &&
              renderTable(
                ["SL", "Upazila", "District"],
                upazilaRows.map((u: any, idx: number) => (
                  <tr
                    key={u.id}
                    className={`${
                      idx % 2 === 0 ? "bg-white" : "bg-slate-100"
                    } hover:bg-emerald-100`}
                  >
                    <td className="border border-slate-300 px-2 py-1.5">
                      {idx + 1}
                    </td>
                    <td className="border border-slate-300 px-2 py-1.5 font-semibold">
                      {u.name}
                    </td>
                    <td className="border border-slate-300 px-2 py-1.5">
                      {districtById[String(u.district_id)]?.name || ""}
                    </td>
                  </tr>
                )),
                "No upazilas found."
              )}

            {activeTab === "union" &&
              renderTable(
                ["SL", "Union", "Upazila", "District"],
                unionRows.map((un: any, idx: number) => {
                  const upa = upazilaById[String(un.upazila_id)];
                  return (
                    <tr
                      key={un.id}
                      className={`${
                        idx % 2 === 0 ? "bg-white" : "bg-slate-100"
                      } hover:bg-emerald-100`}
                    >
                      <td className="border border-slate-300 px-2 py-1.5">
                        {idx + 1}
                      </td>
                      <td className="border border-slate-300 px-2 py-1.5 font-semibold">
                        {un.name}
                      </td>
                      <td className="border border-slate-300 px-2 py-1.5">
                        {upa?.name || ""}
                      </td>
                      <td className="border border-slate-300 px-2 py-1.5">
                        {upa
                          ? districtById[String(upa.district_id)]?.name || ""
                          : ""}
                      </td>
                    </tr>
                  );
                }),
                "No unions found."
              )}

            {activeTab === "bazar" &&
              renderTable(
                ["SL", "Bazar", "Union", "Upazila", "District", "Action"],
                bazarRows.map((b: any, idx: number) => {
                  const un = unionById[String(b.union_id)];
                  const upa = un ? upazilaById[String(un.upazila_id)] : null;
                  return (
                    <tr
                      key={b.id}
                      className={`${
                        idx % 2 === 0 ? "bg-white" : "bg-slate-100"
                      } hover:bg-emerald-100`}
                    >
                      <td className="border border-slate-300 px-2 py-1.5">
                        {idx + 1}
                      </td>
                      <td className="border border-slate-300 px-2 py-1.5 font-semibold">
                        {b.name}
                      </td>
                      <td className="border border-slate-300 px-2 py-1.5">
                        {un?.name || ""}
                      </td>
                      <td className="border border-slate-300 px-2 py-1.5">
                        {upa?.name || ""}
                      </td>
                      <td className="border border-slate-300 px-2 py-1.5">
                        {upa
                          ? districtById[String(upa.district_id)]?.name || ""
                          : ""}
                      </td>
                      <td className="border border-slate-300 px-2 py-1.5 text-center whitespace-nowrap">
                        <button
                          onClick={() => openEditBazar(b)}
                          className="inline-flex items-center px-2 py-0.5 rounded-md bg-blue-600 hover:bg-blue-700 text-white text-[11px] font-bold shadow-sm transition mr-1"
                        >
                          Edit
                        </button>
                        {canDeleteBazar && (
                          <button
                            onClick={() => handleDeleteBazar(b)}
                            className="inline-flex items-center px-2 py-0.5 rounded-md bg-red-600 hover:bg-red-700 text-white text-[11px] font-bold shadow-sm transition"
                          >
                            Delete
                          </button>
                        )}
                      </td>
                    </tr>
                  );
                }),
                "No bazars found."
              )}
          </>
        )}
      </div>

      {/* 3. BAZAR ADD / EDIT MODAL */}
      {showModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-3">
          <div className="bg-white rounded-xl shadow-xl w-full max-w-md max-h-[90vh] overflow-y-auto">
            <div className="flex items-center justify-between px-4 py-2.5 border-b border-slate-300 bg-slate-100">
              <h2 className="text-sm font-bold text-slate-800">
                {editingBazarId ? "Edit Bazar" : "Add Bazar"}
              </h2>
              <button
                onClick={() => setShowModal(false)}
                className="text-slate-500 hover:text-slate-800 text-lg leading-none px-1"
              >
                ×
              </button>
            </div>
            <div className="p-4 space-y-3">
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Bazar Name
                </label>
                <input
                  type="text"
                  value={bazarName}
                  onChange={(e) => setBazarName(e.target.value)}
                  placeholder="Enter bazar name"
                  className="w-full rounded border border-slate-300 px-2.5 py-1.5 text-xs outline-none focus:border-emerald-500"
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  District
                </label>
                <SearchableSelect
                  searchable={false}
                  value={modalDistrict}
                  onChange={(v) => {
                    setModalDistrict(v);
                    setModalUpazila("");
                    setModalUnion("");
                  }}
                  placeholder="Select District"
                  options={visibleDistricts.map((d: any) => ({
                    value: d.name,
                    label: d.name,
                  }))}
                />
              </div>
              <div>
                <label className="block text-xs font-bold text-slate-600 mb-1">
                  Upazila
                </label>
                <SearchableSelect
                  value={modalUpazila}
                  onChange={(v) => {
                    setModalUpazila(v);
                    setModalUnion("");
                  }}
                  placeholder="Select Upazila"
                  options={modalUpazilas.map((u: any) => ({
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
                  value={modalUnion}
                  onChange={setModalUnion}
                  placeholder="Select Union"
                  options={modalUnions.map((un: any) => ({
                    value: String(un.id),
                    label: un.name,
                  }))}
                />
              </div>
              <div className="flex items-center justify-end gap-2 pt-1">
                <button
                  onClick={() => setShowModal(false)}
                  className="px-3 py-1.5 rounded text-xs font-semibold bg-slate-100 hover:bg-slate-200 border border-slate-300 text-slate-700"
                >
                  Cancel
                </button>
                <button
                  onClick={handleSaveBazar}
                  disabled={savingBazar}
                  className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white px-3 py-1.5 rounded text-xs font-semibold"
                >
                  {savingBazar ? "Saving..." : "Save"}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
