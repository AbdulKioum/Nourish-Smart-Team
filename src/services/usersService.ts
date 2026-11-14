import { supabase } from "../lib/supabase";

export interface UserData {
  id?: string;
  email: string;
  password?: string;
  name?: string;
  full_name?: string;
  role: string;
  status?: string;
  districts?: string[];
  upazilas?: string[];
  assigned_categories?: string[];
}

// 1. Fetch all users from public.users table
export const getUsers = async () => {
  const { data, error } = await supabase
    .from("users")
    .select("*")
    .order("created_at", { ascending: false });

  if (error) {
    throw new Error(error.message);
  }
  return data || [];
};

// 2. Create User via Edge Function + Save Assigned Areas directly
export const createUser = async (userData: UserData) => {
  const { data, error } = await supabase.functions.invoke("manage-users", {
    body: { action: "CREATE", userData },
  });

  if (data?.error) {
    throw new Error(data.error);
  }

  if (error) {
    try {
      const errRes = await error.context.json();
      if (errRes?.error) throw new Error(errRes.error);
    } catch {
      // Fallback if context JSON parsing fails
    }
    throw new Error(error.message || "Failed to create user");
  }

  // Edge Function ইউজার তৈরি করার পর public.users টেবিলে districts ও upazilas সেভ করা
  if (
    userData.districts !== undefined ||
    userData.upazilas !== undefined ||
    userData.assigned_categories !== undefined
  ) {
    const { error: areaError } = await supabase
      .from("users")
      .update({
        districts: userData.districts || [],
        upazilas: userData.upazilas || [],
        assigned_categories: userData.assigned_categories || [],
      })
      .eq("email", userData.email);

    if (areaError) {
      console.error("Failed to save assigned areas on create:", areaError);
      throw new Error(areaError.message);
    }
  }

  return data;
};

// 3. Update User via Edge Function + Guarantee Direct Table Update
export const updateUser = async (userId: string, userData: Partial<UserData>) => {
  // Edge function call (optional/auth sync-er jonno)
  try {
    await supabase.functions.invoke("manage-users", {
      body: { action: "UPDATE", userId, userData },
    });
  } catch (e) {
    console.warn("Edge function warning:", e);
  }

  // Sorasori public.users table-e shob data নিখুঁতভাবে save kora
  const updatePayload: Record<string, any> = {};
  
  if (userData.name !== undefined) updatePayload.full_name = userData.name;
  if (userData.full_name !== undefined) updatePayload.full_name = userData.full_name;
  if (userData.role !== undefined) updatePayload.role = userData.role;
  if (userData.status !== undefined) updatePayload.status = userData.status;
  if (userData.districts !== undefined) updatePayload.districts = userData.districts;
  if (userData.upazilas !== undefined) updatePayload.upazilas = userData.upazilas;
  if (userData.assigned_categories !== undefined)
    updatePayload.assigned_categories = userData.assigned_categories;

  const { data, error } = await supabase
    .from("users")
    .update(updatePayload)
    .eq("id", userId)
    .select()
    .single();

  if (error) {
    console.error("Failed to update user in public.users:", error);
    throw new Error(error.message);
  }

  return data;
};

// Alias export for backward compatibility with components using 'updateUserRole'
export { updateUser as updateUserRole };

// 4. Delete User via Edge Function
export const deleteUser = async (userId: string) => {
  const { data, error } = await supabase.functions.invoke("manage-users", {
    body: {
      action: "DELETE",
      userId: userId,
      userData: { id: userId }, // Backup payload
    },
  });

  if (data?.error) {
    throw new Error(data.error);
  }

  if (error) {
    try {
      const errRes = await error.context.json();
      if (errRes?.error) throw new Error(errRes.error);
    } catch {
      // Fallback
    }
    throw new Error(error.message || "Failed to delete user");
  }

  return data;
};
