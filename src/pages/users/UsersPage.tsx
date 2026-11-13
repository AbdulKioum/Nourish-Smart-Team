import React, { useEffect, useState } from "react";
import {
  getUsers,
  createUser,
  updateUser,
  deleteUser,
  updateUserRole,
} from "../../services/usersService";
import { supabase } from "../../lib/supabase";
import {
  FaUserPlus,
  FaEdit,
  FaTrash,
  FaSpinner,
  FaSearch,
  FaTimes,
  FaUsers,
  FaMapMarkerAlt,
} from "react-icons/fa";
import {
  compressImageToTinyBase64,
  UserAvatar,
} from "../../utils/avatarHelper";

interface UserRecord {
  id: string;
  email: string;
  full_name: string;
  role: string;
  status: string;
  designation?: string;
  region?: string;
  avatar_url?: string | null;
  districts?: string[];
  upazilas?: string[];
  created_at?: string;
}

interface DistrictItem {
  id: string | number;
  name: string;
}

interface UpazilaItem {
  id: string | number;
  name: string;
  district_id?: string | number;
  district_name?: string;
  district?: string;
}

export default function UsersPage() {
  const [users, setUsers] = useState<UserRecord[]>([]);
  const [loading, setLoading] = useState(true);
  const [submitting, setSubmitting] = useState(false);
  const [deletingId, setDeletingId] = useState<string | null>(null);
  const [errorMsg, setErrorMsg] = useState("");
  const [successMsg, setSuccessMsg] = useState("");
  const [currentUserId, setCurrentUserId] = useState<string | null>(null);
  const [currentEmail, setCurrentEmail] = useState("");

  const [districtsList, setDistrictsList] = useState<DistrictItem[]>([]);
  const [upazilasList, setUpazilasList] = useState<UpazilaItem[]>([]);
  const [loadingLocations, setLoadingLocations] = useState(false);

  const [isDistrictDropdownOpen, setIsDistrictDropdownOpen] = useState(false);
  const [isUpazilaDropdownOpen, setIsUpazilaDropdownOpen] = useState(false);

  const [search, setSearch] = useState("");
  const [roleFilter, setRoleFilter] = useState("");
  const [statusFilter, setStatusFilter] = useState("");

  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingUser, setEditingUser] = useState<UserRecord | null>(null);

  const [formData, setFormData] = useState<{
    email: string;
    password: string;
    full_name: string;
    designation: string;
    region: string;
    avatar_url: string;
    role: string;
    status: string;
    districts: string[];
    upazilas: string[];
  }>({
    email: "",
    password: "",
    full_name: "",
    designation: "Customer Service Officer",
    region: "Mymensingh Zone",
    avatar_url: "",
    role: "officer",
    status: "active",
    districts: [],
    upazilas: [],
  });

  useEffect(() => {
    loadCurrentUser();
    loadUsers();
    loadLocations();
  }, []);

  async function loadCurrentUser() {
    const {
      data: { user },
    } = await supabase.auth.getUser();

    setCurrentUserId(user?.id || null);
    setCurrentEmail(user?.email || "");
  }

  async function loadLocations() {
    try {
      setLoadingLocations(true);

      const { data: distData, error: distError } = await supabase
        .from("districts")
        .select("*")
        .order("name", { ascending: true });

      if (distError) throw distError;

      const { data: upaData, error: upaError } = await supabase
        .from("upazilas")
        .select("*")
        .order("name", { ascending: true });

      if (upaError) throw upaError;

      setDistrictsList(distData || []);
      setUpazilasList(upaData || []);
    } catch (error: any) {
      console.error("Failed to load locations:", error?.message);
    } finally {
      setLoadingLocations(false);
    }
  }

  async function loadUsers() {
    try {
      setLoading(true);
      setErrorMsg("");

      const data = await getUsers();
      const { data: avatarsData } = await supabase
        .from("users")
        .select("id, avatar_url");

      const avatarMap = new Map<string, string | null>();
      (avatarsData || []).forEach((row: any) => {
        avatarMap.set(String(row.id), row.avatar_url || null);
      });

      const merged = (data || []).map((u: any) => ({
        ...u,
        avatar_url: u.avatar_url ?? avatarMap.get(String(u.id)) ?? null,
      }));

      setUsers(merged);
    } catch (error: any) {
      setErrorMsg(error?.message || "Failed to load users.");
    } finally {
      setLoading(false);
    }
  }

  async function handlePhotoSelect(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const tinyBase64 = await compressImageToTinyBase64(file, 80);
      setFormData((prev) => ({ ...prev, avatar_url: tinyBase64 }));
    } catch (err) {
      console.error("Failed to compress image:", err);
    }
  }

  function parseArrayField(field: any): string[] {
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

  const availableUpazilas = upazilasList.filter((upazila) => {
    if (formData.districts.length === 0) return false;

    const matchedDistrictById = districtsList.find(
      (d) => String(d.id) === String(upazila.district_id)
    );
    if (
      matchedDistrictById &&
      formData.districts.includes(matchedDistrictById.name)
    ) {
      return true;
    }

    const distName = upazila.district_name || upazila.district;
    if (distName && formData.districts.includes(distName)) {
      return true;
    }

    return false;
  });

  function handleToggleDistrict(districtName: string) {
    setFormData((prev) => {
      const isSelected = prev.districts.includes(districtName);
      const updatedDistricts = isSelected
        ? prev.districts.filter((d) => d !== districtName)
        : [...prev.districts, districtName];

      const validUpazilaNames = upazilasList
        .filter((u) => {
          const matchedDist = districtsList.find(
            (d) => String(d.id) === String(u.district_id)
          );
          const dName = matchedDist?.name || u.district_name || u.district;
          return dName && updatedDistricts.includes(dName);
        })
        .map((u) => u.name);

      const updatedUpazilas = prev.upazilas.filter((uName) =>
        validUpazilaNames.includes(uName)
      );

      return {
        ...prev,
        districts: updatedDistricts,
        upazilas: updatedUpazilas,
      };
    });
  }

  function handleToggleUpazila(upazilaName: string) {
    setFormData((prev) => {
      const isSelected = prev.upazilas.includes(upazilaName);
      const updatedUpazilas = isSelected
        ? prev.upazilas.filter((u) => u !== upazilaName)
        : [...prev.upazilas, upazilaName];

      return {
        ...prev,
        upazilas: updatedUpazilas,
      };
    });
  }

  function handleSelectAllUpazilas() {
    const allNames = availableUpazilas.map((u) => u.name);
    const allSelected = allNames.every((name) =>
      formData.upazilas.includes(name)
    );

    setFormData((prev) => ({
      ...prev,
      upazilas: allSelected ? [] : allNames,
    }));
  }

  function handleOpenCreateModal() {
    setEditingUser(null);
    setErrorMsg("");
    setSuccessMsg("");
    setIsDistrictDropdownOpen(false);
    setIsUpazilaDropdownOpen(false);

    setFormData({
      email: "",
      password: "",
      full_name: "",
      designation: "Customer Service Officer",
      region: "Mymensingh Zone",
      avatar_url: "",
      role: "officer",
      status: "active",
      districts: [],
      upazilas: [],
    });

    setIsModalOpen(true);
  }

  function handleOpenEditModal(user: UserRecord) {
    setEditingUser(user);
    setErrorMsg("");
    setSuccessMsg("");
    setIsDistrictDropdownOpen(false);
    setIsUpazilaDropdownOpen(false);

    const allowedRoles = ["super_admin", "manager", "officer"];
    const validRole = allowedRoles.includes(user.role) ? user.role : "officer";

    setFormData({
      email: user.email || "",
      password: "",
      full_name: user.full_name || "",
      designation: user.designation || "Customer Service Officer",
      region: user.region || "Mymensingh Zone",
      avatar_url: user.avatar_url || "",
      role: validRole,
      status: user.status || "active",
      districts: parseArrayField(user.districts),
      upazilas: parseArrayField(user.upazilas),
    });

    setIsModalOpen(true);
  }

  function closeModal() {
    if (submitting) return;
    setIsModalOpen(false);
    setEditingUser(null);
    setErrorMsg("");
    setIsDistrictDropdownOpen(false);
    setIsUpazilaDropdownOpen(false);
  }

  async function handleSave(e: React.FormEvent) {
    e.preventDefault();
    setSubmitting(true);
    setErrorMsg("");
    setSuccessMsg("");

    try {
      if (editingUser) {
        await updateUser(editingUser.id, {
          name: formData.full_name.trim(),
          role: formData.role,
          status: formData.status,
          districts: formData.districts,
          upazilas: formData.upazilas,
        });

        // Also directly update designation, region, full_name, avatar_url in `users` table
        await supabase
          .from("users")
          .update({
            full_name: formData.full_name.trim(),
            designation: formData.designation.trim(),
            region: formData.region.trim(),
            avatar_url: formData.avatar_url || null,
            role: formData.role,
            status: formData.status,
            districts: formData.districts,
            upazilas: formData.upazilas,
          })
          .eq("id", editingUser.id);

        setSuccessMsg("User updated successfully.");
      } else {
        if (!formData.full_name.trim())
          throw new Error("Full name is required.");
        if (!formData.email.trim()) throw new Error("Email is required.");
        if (!formData.password) throw new Error("Password is required.");
        if (formData.password.length < 6)
          throw new Error("Password must be at least 6 characters long.");

        const created: any = await createUser({
          name: formData.full_name.trim(),
          email: formData.email.trim(),
          password: formData.password,
          role: formData.role,
          status: formData.status,
          districts: formData.districts,
          upazilas: formData.upazilas,
        });

        const newId = created?.id || created?.user?.id;
        if (newId) {
          await supabase
            .from("users")
            .update({
              designation: formData.designation.trim(),
              region: formData.region.trim(),
              avatar_url: formData.avatar_url || null,
            })
            .eq("id", newId);
        } else {
          await supabase
            .from("users")
            .update({
              designation: formData.designation.trim(),
              region: formData.region.trim(),
              avatar_url: formData.avatar_url || null,
            })
            .eq("email", formData.email.trim());
        }

        setSuccessMsg("User created successfully.");
      }

      await loadUsers();

      setTimeout(() => {
        setIsModalOpen(false);
        setSuccessMsg("");
      }, 500);
    } catch (error: any) {
      setErrorMsg(error?.message || "Something went wrong.");
    } finally {
      setSubmitting(false);
    }
  }

  async function handleDelete(user: UserRecord) {
    if (user.id === currentUserId || user.email === currentEmail) {
      alert("You cannot delete your own account.");
      return;
    }

    const confirmed = window.confirm(
      `Are you sure you want to delete "${user.full_name || user.email}"?`
    );
    if (!confirmed) return;

    try {
      setDeletingId(user.id);
      setErrorMsg("");
      setSuccessMsg("");
      await deleteUser(user.id);
      setSuccessMsg("User deleted successfully.");
      await loadUsers();
    } catch (error: any) {
      setErrorMsg(error?.message || "Failed to delete user.");
    } finally {
      setDeletingId(null);
    }
  }

  async function handleRoleChange(id: string, role: string) {
    try {
      await updateUserRole(id, { role });
      await loadUsers();
    } catch (error: any) {
      setErrorMsg(error?.message || "Failed to update role.");
    }
  }

  async function handleStatusChange(id: string, status: string) {
    try {
      await updateUser(id, { status });
      await loadUsers();
    } catch (error: any) {
      setErrorMsg(error?.message || "Failed to update status.");
    }
  }

  const filteredUsers = users.filter((user) => {
    const searchText = search.trim().toLowerCase();
    const userDistricts = parseArrayField(user.districts)
      .join(" ")
      .toLowerCase();
    const userUpazilas = parseArrayField(user.upazilas)
      .join(" ")
      .toLowerCase();

    const matchSearch =
      !searchText ||
      user.full_name?.toLowerCase().includes(searchText) ||
      user.email?.toLowerCase().includes(searchText) ||
      user.designation?.toLowerCase().includes(searchText) ||
      user.region?.toLowerCase().includes(searchText) ||
      userDistricts.includes(searchText) ||
      userUpazilas.includes(searchText);

    const matchRole = !roleFilter || user.role === roleFilter;
    const matchStatus = !statusFilter || user.status === statusFilter;

    return matchSearch && matchRole && matchStatus;
  });

  function getRoleLabel(role: string) {
    switch (role) {
      case "super_admin":
        return "Super Admin";
      case "manager":
        return "Manager";
      case "officer":
        return "Officer";
      default:
        return (
          role
            ?.replaceAll("_", " ")
            .replace(/\b\w/g, (char) => char.toUpperCase()) || "Officer"
        );
    }
  }

  function getRoleBadge(role: string) {
    switch (role) {
      case "super_admin":
        return "bg-purple-100 text-purple-700";
      case "manager":
        return "bg-blue-100 text-blue-700";
      case "officer":
        return "bg-cyan-100 text-cyan-700";
      default:
        return "bg-slate-100 text-slate-700";
    }
  }

  const activeCount = users.filter((u) => u.status === "active").length;

  return (
    <div className="p-3 sm:p-4 md:p-6 max-w-7xl mx-auto overflow-x-hidden">
      {/* HEADER */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 mb-4">
        <div className="flex flex-wrap items-center gap-2.5 sm:gap-3">
          <div className="w-9 h-9 sm:w-10 sm:h-10 shrink-0 rounded-xl bg-emerald-100 text-emerald-600 flex items-center justify-center">
            <FaUsers className="text-base sm:text-lg" />
          </div>

          <h1 className="text-lg sm:text-xl font-bold text-slate-800">
            User Management
          </h1>

          <div className="flex items-center gap-1.5 ml-0 sm:ml-2">
            <span className="px-2.5 py-1 rounded-lg bg-slate-100 border border-slate-200 text-slate-700 text-xs font-medium">
              Total: <strong className="font-bold">{users.length}</strong>
            </span>

            <span className="px-2.5 py-1 rounded-lg bg-emerald-50 border border-emerald-200 text-emerald-700 text-xs font-medium">
              Active: <strong className="font-bold">{activeCount}</strong>
            </span>

            <span className="px-2.5 py-1 rounded-lg bg-blue-50 border border-blue-200 text-blue-700 text-xs font-medium">
              View: <strong className="font-bold">{filteredUsers.length}</strong>
            </span>
          </div>
        </div>

        <button
          onClick={handleOpenCreateModal}
          className="w-full sm:w-auto flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 active:scale-[0.99] text-white px-4 py-2 rounded-xl text-xs sm:text-sm font-medium shadow-sm transition shrink-0"
        >
          <FaUserPlus />
          Add New User
        </button>
      </div>

      {/* ALERTS */}
      {errorMsg && (
        <div className="mb-3 flex items-center justify-between gap-3 bg-red-50 border border-red-200 text-red-700 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm">
          <span className="break-words">{errorMsg}</span>
          <button
            onClick={() => setErrorMsg("")}
            className="text-red-500 hover:text-red-700 shrink-0 p-1"
          >
            <FaTimes />
          </button>
        </div>
      )}

      {successMsg && (
        <div className="mb-3 bg-emerald-50 border border-emerald-200 text-emerald-700 px-3.5 py-2.5 rounded-xl text-xs sm:text-sm">
          {successMsg}
        </div>
      )}

      {/* SEARCH / FILTER */}
      <div className="bg-white border border-slate-200 rounded-2xl p-3 mb-4 shadow-sm">
        <div className="grid grid-cols-2 md:grid-cols-3 gap-2.5">
          <div className="relative col-span-2 md:col-span-1">
            <FaSearch className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 text-xs sm:text-sm" />
            <input
              type="text"
              placeholder="Search name, email, designation, area..."
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full border border-slate-300 rounded-xl pl-8 pr-3 py-2 text-xs sm:text-sm outline-none focus:ring-2 focus:ring-emerald-200 focus:border-emerald-500"
            />
          </div>

          <select
            value={roleFilter}
            onChange={(e) => setRoleFilter(e.target.value)}
            className="border border-slate-300 rounded-xl px-2.5 py-2 text-xs sm:text-sm outline-none bg-white focus:ring-2 focus:ring-emerald-200 focus:border-emerald-500"
          >
            <option value="">All Roles</option>
            <option value="super_admin">Super Admin</option>
            <option value="manager">Manager</option>
            <option value="officer">Officer</option>
          </select>

          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="border border-slate-300 rounded-xl px-2.5 py-2 text-xs sm:text-sm outline-none bg-white focus:ring-2 focus:ring-emerald-200 focus:border-emerald-500"
          >
            <option value="">All Status</option>
            <option value="active">Active</option>
            <option value="inactive">Inactive</option>
          </select>
        </div>
      </div>

      {/* MOBILE CARD VIEW */}
      <div className="block md:hidden space-y-3">
        {loading ? (
          <div className="bg-white border border-slate-200 rounded-2xl py-10 text-center text-slate-500 text-sm">
            <FaSpinner className="animate-spin inline mr-2" />
            Loading users...
          </div>
        ) : filteredUsers.length === 0 ? (
          <div className="bg-white border border-slate-200 rounded-2xl py-10 text-center text-slate-400 text-sm">
            No users found.
          </div>
        ) : (
          filteredUsers.map((user) => {
            const isCurrentUser =
              user.id === currentUserId || user.email === currentEmail;
            const userDistricts = parseArrayField(user.districts);
            const userUpazilas = parseArrayField(user.upazilas);

            return (
              <div
                key={user.id}
                className="bg-white border border-slate-200 rounded-2xl p-3.5 shadow-sm space-y-3"
              >
                <div className="flex items-start justify-between gap-2">
                  <div className="flex items-center gap-2.5 min-w-0">
                    <UserAvatar
                      name={user.full_name || user.email || "U"}
                      avatarUrl={user.avatar_url}
                      size="w-9 h-9"
                      textSize="text-sm"
                    />
                    <div className="min-w-0">
                      <div className="font-semibold text-slate-800 text-sm truncate flex items-center gap-1.5">
                        <span className="truncate">
                          {user.full_name || "Unnamed User"}
                        </span>
                        {isCurrentUser && (
                          <span className="shrink-0 text-[10px] bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded-full">
                            You
                          </span>
                        )}
                      </div>
                      <div className="text-xs text-slate-500 truncate">
                        {user.email}
                      </div>
                      {(user.designation || user.region) && (
                        <div className="text-[11px] text-emerald-700 font-medium truncate">
                          {[user.designation, user.region]
                            .filter(Boolean)
                            .join(" • ")}
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="flex flex-col items-end gap-1 shrink-0">
                    <span
                      className={`px-2 py-0.5 rounded-full text-[10px] font-semibold ${getRoleBadge(
                        user.role
                      )}`}
                    >
                      {getRoleLabel(user.role)}
                    </span>
                    <span
                      className={`px-2 py-0.5 rounded-full text-[10px] font-semibold ${
                        user.status === "active"
                          ? "bg-emerald-100 text-emerald-700"
                          : "bg-red-100 text-red-700"
                      }`}
                    >
                      {user.status === "active" ? "Active" : "Inactive"}
                    </span>
                  </div>
                </div>

                <div className="bg-slate-50 rounded-xl p-2.5 border border-slate-100">
                  <p className="text-[10px] font-semibold text-slate-400 uppercase mb-1">
                    Assigned Area
                  </p>
                  {userDistricts.length === 0 && userUpazilas.length === 0 ? (
                    <span className="text-xs text-slate-400 italic">
                      Not assigned
                    </span>
                  ) : (
                    <div className="space-y-1.5">
                      {userDistricts.length > 0 && (
                        <div className="flex flex-wrap gap-1">
                          {userDistricts.map((d) => (
                            <span
                              key={d}
                              className="px-2 py-0.5 bg-blue-50 text-blue-700 border border-blue-200 rounded-md text-[11px] font-medium"
                            >
                              {d}
                            </span>
                          ))}
                        </div>
                      )}
                      {userUpazilas.length > 0 && (
                        <div className="flex flex-wrap gap-1">
                          {userUpazilas.map((u) => (
                            <span
                              key={u}
                              className="px-2 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-md text-[11px]"
                            >
                              {u}
                            </span>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </div>

                <div className="flex items-center justify-end gap-2 pt-2 border-t border-slate-100">
                  <button
                    onClick={() => handleOpenEditModal(user)}
                    className="flex-1 inline-flex items-center justify-center gap-1.5 py-2 rounded-xl bg-blue-50 text-blue-600 active:bg-blue-100 text-xs font-semibold transition"
                  >
                    <FaEdit />
                    Edit User & Area
                  </button>

                  <button
                    onClick={() => handleDelete(user)}
                    disabled={isCurrentUser || deletingId === user.id}
                    className="inline-flex items-center justify-center gap-1.5 px-3.5 py-2 rounded-xl bg-red-50 text-red-600 active:bg-red-100 disabled:opacity-40 text-xs font-semibold transition"
                  >
                    {deletingId === user.id ? (
                      <FaSpinner className="animate-spin" />
                    ) : (
                      <FaTrash />
                    )}
                    Delete
                  </button>
                </div>
              </div>
            );
          })
        )}
      </div>

      {/* DESKTOP TABLE VIEW */}
      <div className="hidden md:block bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="min-w-full">
            <thead className="bg-slate-50 border-b border-slate-200">
              <tr>
                <th className="px-5 py-3 text-left text-xs font-semibold text-slate-500 uppercase">
                  User & Designation
                </th>
                <th className="px-5 py-3 text-left text-xs font-semibold text-slate-500 uppercase">
                  Assigned Area
                </th>
                <th className="px-5 py-3 text-left text-xs font-semibold text-slate-500 uppercase">
                  Role
                </th>
                <th className="px-5 py-3 text-left text-xs font-semibold text-slate-500 uppercase">
                  Status
                </th>
                <th className="px-5 py-3 text-right text-xs font-semibold text-slate-500 uppercase">
                  Actions
                </th>
              </tr>
            </thead>

            <tbody className="divide-y divide-slate-100">
              {loading ? (
                <tr>
                  <td colSpan={5} className="py-12 text-center text-slate-500">
                    <FaSpinner className="animate-spin inline mr-2" />
                    Loading users...
                  </td>
                </tr>
              ) : filteredUsers.length === 0 ? (
                <tr>
                  <td colSpan={5} className="py-12 text-center">
                    <div className="text-slate-400 text-sm">
                      No users found.
                    </div>
                  </td>
                </tr>
              ) : (
                filteredUsers.map((user) => {
                  const isCurrentUser =
                    user.id === currentUserId || user.email === currentEmail;

                  const userDistricts = parseArrayField(user.districts);
                  const userUpazilas = parseArrayField(user.upazilas);

                  return (
                    <tr key={user.id} className="hover:bg-slate-50 transition">
                      <td className="px-5 py-4">
                        <div className="flex items-center gap-3">
                          <UserAvatar
                            name={user.full_name || user.email || "U"}
                            avatarUrl={user.avatar_url}
                            size="w-10 h-10"
                            textSize="text-sm"
                          />

                          <div>
                            <div className="font-semibold text-slate-800">
                              {user.full_name || "Unnamed User"}
                              {isCurrentUser && (
                                <span className="ml-2 text-[10px] bg-blue-100 text-blue-700 px-2 py-0.5 rounded-full">
                                  You
                                </span>
                              )}
                            </div>
                            <div className="text-xs text-slate-500">
                              {user.email}
                            </div>
                            {(user.designation || user.region) && (
                              <div className="text-xs text-emerald-700 font-medium mt-0.5">
                                {[user.designation, user.region]
                                  .filter(Boolean)
                                  .join(" • ")}
                              </div>
                            )}
                          </div>
                        </div>
                      </td>

                      <td className="px-5 py-4">
                        {userDistricts.length === 0 &&
                        userUpazilas.length === 0 ? (
                          <span className="text-xs text-slate-400 italic">
                            Not assigned
                          </span>
                        ) : (
                          <div className="space-y-1.5 max-w-xs">
                            {userDistricts.length > 0 && (
                              <div className="flex flex-wrap gap-1">
                                {userDistricts.map((d) => (
                                  <span
                                    key={d}
                                    className="px-2 py-0.5 bg-blue-50 text-blue-700 border border-blue-200 rounded-md text-[11px] font-medium"
                                  >
                                    {d}
                                  </span>
                                ))}
                              </div>
                            )}

                            {userUpazilas.length > 0 && (
                              <div className="flex flex-wrap gap-1">
                                {userUpazilas.map((u) => (
                                  <span
                                    key={u}
                                    className="px-2 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded-md text-[11px]"
                                  >
                                    {u}
                                  </span>
                                ))}
                              </div>
                            )}
                          </div>
                        )}
                      </td>

                      <td className="px-5 py-4">
                        <div className="flex items-center gap-2">
                          <span
                            className={`px-2.5 py-1 rounded-full text-xs font-semibold ${getRoleBadge(
                              user.role
                            )}`}
                          >
                            {getRoleLabel(user.role)}
                          </span>

                          {!isCurrentUser && (
                            <select
                              value={user.role}
                              onChange={(e) =>
                                handleRoleChange(user.id, e.target.value)
                              }
                              className="border border-slate-200 rounded-lg px-2 py-1 text-xs bg-white"
                            >
                              <option value="super_admin">Super Admin</option>
                              <option value="manager">Manager</option>
                              <option value="officer">Officer</option>
                            </select>
                          )}
                        </div>
                      </td>

                      <td className="px-5 py-4">
                        <div className="flex items-center gap-2">
                          <span
                            className={`px-2.5 py-1 rounded-full text-xs font-semibold ${
                              user.status === "active"
                                ? "bg-emerald-100 text-emerald-700"
                                : "bg-red-100 text-red-700"
                            }`}
                          >
                            {user.status === "active" ? "Active" : "Inactive"}
                          </span>

                          {!isCurrentUser && (
                            <select
                              value={user.status}
                              onChange={(e) =>
                                handleStatusChange(user.id, e.target.value)
                              }
                              className="border border-slate-200 rounded-lg px-2 py-1 text-xs bg-white"
                            >
                              <option value="active">Active</option>
                              <option value="inactive">Inactive</option>
                            </select>
                          )}
                        </div>
                      </td>

                      <td className="px-5 py-4 text-right">
                        <div className="flex justify-end gap-2">
                          <button
                            onClick={() => handleOpenEditModal(user)}
                            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-blue-50 text-blue-600 hover:bg-blue-100 text-sm font-medium transition"
                          >
                            <FaEdit />
                            Edit
                          </button>

                          <button
                            onClick={() => handleDelete(user)}
                            disabled={isCurrentUser || deletingId === user.id}
                            className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-red-50 text-red-600 hover:bg-red-100 disabled:opacity-40 disabled:cursor-not-allowed text-sm font-medium transition"
                          >
                            {deletingId === user.id ? (
                              <FaSpinner className="animate-spin" />
                            ) : (
                              <FaTrash />
                            )}
                            Delete
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* CREATE / EDIT MODAL */}
      {isModalOpen && (
        <div className="fixed inset-0 z-[100] bg-black/50 backdrop-blur-sm flex items-end sm:items-center justify-center p-0 sm:p-4">
          <div className="bg-white rounded-t-2xl sm:rounded-2xl shadow-2xl w-full max-w-lg max-h-[92vh] sm:max-h-[90vh] flex flex-col overflow-hidden">
            <div className="flex items-center justify-between px-4 sm:px-6 py-3.5 sm:py-4 border-b border-slate-200 bg-white shrink-0">
              <h2 className="text-base sm:text-lg font-bold text-slate-800">
                {editingUser ? "Edit User Details & Area" : "Add New User"}
              </h2>

              <button
                type="button"
                onClick={closeModal}
                disabled={submitting}
                className="w-8 h-8 sm:w-9 sm:h-9 rounded-lg hover:bg-slate-100 flex items-center justify-center text-slate-500"
              >
                <FaTimes />
              </button>
            </div>

            <form
              onSubmit={handleSave}
              className="p-4 sm:p-6 space-y-4 overflow-y-auto flex-1"
            >
              {/* PROFILE PHOTO UPLOAD */}
              <div className="flex items-center gap-3 p-3 rounded-xl bg-slate-50 border border-slate-200">
                <UserAvatar
                  name={formData.full_name || formData.email || "U"}
                  avatarUrl={formData.avatar_url}
                  size="w-11 h-11"
                  textSize="text-base"
                />
                <div className="flex-1 min-w-0">
                  <div className="text-xs font-semibold text-slate-700">
                    Profile Photo
                  </div>
                  
                </div>
                <div className="flex items-center gap-1.5 shrink-0">
                  <label className="cursor-pointer px-2.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold transition">
                    {formData.avatar_url ? "Change" : "Upload"}
                    <input
                      type="file"
                      accept="image/*"
                      onChange={handlePhotoSelect}
                      className="hidden"
                    />
                  </label>
                  {formData.avatar_url && (
                    <button
                      type="button"
                      onClick={() =>
                        setFormData((prev) => ({ ...prev, avatar_url: "" }))
                      }
                      className="px-2 py-1.5 rounded-lg border border-slate-300 bg-white hover:bg-slate-100 text-slate-600 text-xs font-semibold"
                    >
                      Remove
                    </button>
                  )}
                </div>
              </div>

              <div>
                <label className="block text-xs sm:text-sm font-medium text-slate-700 mb-1">
                  Full Name *
                </label>
                <input
                  type="text"
                  required
                  value={formData.full_name}
                  onChange={(e) =>
                    setFormData({ ...formData, full_name: e.target.value })
                  }
                  placeholder="Enter full name"
                  className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none focus:border-emerald-500"
                />
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs sm:text-sm font-medium text-slate-700 mb-1">
                    Designation
                  </label>
                  <input
                    type="text"
                    value={formData.designation}
                    onChange={(e) =>
                      setFormData({ ...formData, designation: e.target.value })
                    }
                    placeholder="e.g. Customer Service Officer"
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none focus:border-emerald-500"
                  />
                </div>

                <div>
                  <label className="block text-xs sm:text-sm font-medium text-slate-700 mb-1">
                    Region / Zone
                  </label>
                  <input
                    type="text"
                    value={formData.region}
                    onChange={(e) =>
                      setFormData({ ...formData, region: e.target.value })
                    }
                    placeholder="e.g. Mymensingh Zone"
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none focus:border-emerald-500"
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs sm:text-sm font-medium text-slate-700 mb-1">
                  Email
                </label>
                <input
                  type="email"
                  required
                  disabled={!!editingUser}
                  value={formData.email}
                  onChange={(e) =>
                    setFormData({ ...formData, email: e.target.value })
                  }
                  placeholder="user@example.com"
                  className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none disabled:bg-slate-100 disabled:text-slate-500"
                />
              </div>

              {!editingUser && (
                <div>
                  <label className="block text-xs sm:text-sm font-medium text-slate-700 mb-1">
                    Password *
                  </label>
                  <input
                    type="password"
                    required
                    minLength={6}
                    value={formData.password}
                    onChange={(e) =>
                      setFormData({ ...formData, password: e.target.value })
                    }
                    placeholder="Minimum 6 characters"
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none focus:border-emerald-500"
                  />
                </div>
              )}

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs sm:text-sm font-medium text-slate-700 mb-1">
                    Role
                  </label>
                  <select
                    value={formData.role}
                    onChange={(e) =>
                      setFormData({ ...formData, role: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none bg-white"
                  >
                    <option value="super_admin">Super Admin</option>
                    <option value="manager">Manager</option>
                    <option value="officer">Officer</option>
                  </select>
                </div>

                <div>
                  <label className="block text-xs sm:text-sm font-medium text-slate-700 mb-1">
                    Status
                  </label>
                  <select
                    value={formData.status}
                    onChange={(e) =>
                      setFormData({ ...formData, status: e.target.value })
                    }
                    className="w-full border border-slate-300 rounded-xl px-3 py-2 text-sm outline-none bg-white"
                  >
                    <option value="active">Active</option>
                    <option value="inactive">Inactive</option>
                  </select>
                </div>
              </div>

              {/* MULTI-SELECT DISTRICTS */}
              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="block text-xs sm:text-sm font-medium text-slate-700">
                    Assign Districts (Multiple)
                  </label>
                  {formData.districts.length > 0 && (
                    <button
                      type="button"
                      onClick={() =>
                        setFormData({
                          ...formData,
                          districts: [],
                          upazilas: [],
                        })
                      }
                      className="text-xs text-red-500 hover:underline"
                    >
                      Clear All
                    </button>
                  )}
                </div>

                {formData.districts.length > 0 && (
                  <div className="flex flex-wrap gap-1.5 mb-2">
                    {formData.districts.map((dist) => (
                      <span
                        key={dist}
                        className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-blue-50 text-blue-700 border border-blue-200 text-xs font-medium"
                      >
                        <FaMapMarkerAlt className="text-[10px]" />
                        {dist}
                        <button
                          type="button"
                          onClick={() => handleToggleDistrict(dist)}
                          className="ml-1 text-blue-500 hover:text-blue-800"
                        >
                          <FaTimes />
                        </button>
                      </span>
                    ))}
                  </div>
                )}

                <button
                  type="button"
                  onClick={() =>
                    setIsDistrictDropdownOpen(!isDistrictDropdownOpen)
                  }
                  className="w-full flex items-center justify-between border border-slate-300 rounded-xl px-3 py-2 text-sm text-left bg-white hover:bg-slate-50"
                >
                  <span className="text-slate-600">
                    {formData.districts.length === 0
                      ? "Select District(s)..."
                      : `${formData.districts.length} District(s) Selected`}
                  </span>
                  <span className="text-xs text-slate-400">
                    {isDistrictDropdownOpen ? "▲" : "▼"}
                  </span>
                </button>

                {isDistrictDropdownOpen && (
                  <div className="mt-1 border border-slate-200 rounded-xl max-h-40 overflow-y-auto bg-slate-50 p-2 grid grid-cols-1 sm:grid-cols-2 gap-1">
                    {loadingLocations ? (
                      <div className="col-span-2 text-xs text-slate-400 py-2 text-center">
                        Loading districts...
                      </div>
                    ) : (
                      districtsList.map((district) => {
                        const checked = formData.districts.includes(
                          district.name
                        );
                        return (
                          <label
                            key={district.id}
                            className={`flex items-center gap-2.5 px-2.5 py-2 rounded-lg cursor-pointer text-xs sm:text-sm transition ${
                              checked
                                ? "bg-blue-50 text-blue-700 font-medium"
                                : "hover:bg-white text-slate-700"
                            }`}
                          >
                            <input
                              type="checkbox"
                              checked={checked}
                              onChange={() =>
                                handleToggleDistrict(district.name)
                              }
                              className="w-4 h-4 rounded border-slate-300 text-emerald-600"
                            />
                            <span className="truncate">{district.name}</span>
                          </label>
                        );
                      })
                    )}
                  </div>
                )}
              </div>

              {/* MULTI-SELECT UPAZILAS */}
              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="block text-xs sm:text-sm font-medium text-slate-700">
                    Assign Upazilas (Multiple)
                  </label>

                  {availableUpazilas.length > 0 && (
                    <button
                      type="button"
                      onClick={handleSelectAllUpazilas}
                      className="text-xs text-emerald-600 hover:underline font-medium"
                    >
                      {availableUpazilas.every((u) =>
                        formData.upazilas.includes(u.name)
                      )
                        ? "Deselect All"
                        : "Select All"}
                    </button>
                  )}
                </div>

                {formData.upazilas.length > 0 && (
                  <div className="flex flex-wrap gap-1.5 mb-2">
                    {formData.upazilas.map((upa) => (
                      <span
                        key={upa}
                        className="inline-flex items-center gap-1 px-2.5 py-1 rounded-lg bg-emerald-50 text-emerald-700 border border-emerald-200 text-xs font-medium"
                      >
                        {upa}
                        <button
                          type="button"
                          onClick={() => handleToggleUpazila(upa)}
                          className="ml-1 text-emerald-500 hover:text-emerald-800"
                        >
                          <FaTimes />
                        </button>
                      </span>
                    ))}
                  </div>
                )}

                <button
                  type="button"
                  disabled={formData.districts.length === 0}
                  onClick={() =>
                    setIsUpazilaDropdownOpen(!isUpazilaDropdownOpen)
                  }
                  className="w-full flex items-center justify-between border border-slate-300 rounded-xl px-3 py-2 text-sm text-left bg-white hover:bg-slate-50 disabled:bg-slate-100"
                >
                  <span className="text-slate-600">
                    {formData.districts.length === 0
                      ? "Select district first..."
                      : formData.upazilas.length === 0
                      ? "Select Upazila(s)..."
                      : `${formData.upazilas.length} Upazila(s) Selected`}
                  </span>
                  <span className="text-xs text-slate-400">
                    {isUpazilaDropdownOpen ? "▲" : "▼"}
                  </span>
                </button>

                {isUpazilaDropdownOpen && formData.districts.length > 0 && (
                  <div className="mt-1 border border-slate-200 rounded-xl max-h-44 overflow-y-auto bg-slate-50 p-2 grid grid-cols-1 sm:grid-cols-2 gap-1">
                    {availableUpazilas.map((upazila) => {
                      const checked = formData.upazilas.includes(upazila.name);
                      return (
                        <label
                          key={upazila.id}
                          className={`flex items-center gap-2.5 px-2.5 py-2 rounded-lg cursor-pointer text-xs sm:text-sm transition ${
                            checked
                              ? "bg-emerald-50 text-emerald-700 font-medium"
                              : "hover:bg-white text-slate-700"
                          }`}
                        >
                          <input
                            type="checkbox"
                            checked={checked}
                            onChange={() => handleToggleUpazila(upazila.name)}
                            className="w-4 h-4 rounded border-slate-300 text-emerald-600"
                          />
                          <span className="truncate">{upazila.name}</span>
                        </label>
                      );
                    })}
                  </div>
                )}
              </div>

              <div className="flex gap-2.5 pt-3 border-t border-slate-100 sticky bottom-0 bg-white">
                <button
                  type="button"
                  onClick={closeModal}
                  disabled={submitting}
                  className="flex-1 py-2.5 rounded-xl border border-slate-300 text-slate-600 hover:bg-slate-50 text-sm font-medium"
                >
                  Cancel
                </button>

                <button
                  type="submit"
                  disabled={submitting}
                  className="flex-1 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white text-sm font-medium flex items-center justify-center gap-2"
                >
                  {submitting && <FaSpinner className="animate-spin" />}
                  {editingUser ? "Save Changes" : "Create User"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}