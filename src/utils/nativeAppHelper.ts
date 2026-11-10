import { Capacitor } from "@capacitor/core";
import { LocalNotifications } from "@capacitor/local-notifications";
import { Geolocation } from "@capacitor/geolocation";
import { Filesystem, Directory } from "@capacitor/filesystem";
import { Share } from "@capacitor/share";

// ১. অ্যাপ ওপেন হলে নোটিফিকেশন চ্যানেল ও পারমিশন রেডি করা
export async function initNativePermissions() {
  if (!Capacitor.isNativePlatform()) return;
  try {
    const notifPerm = await LocalNotifications.checkPermissions();
    if (notifPerm.display !== "granted") {
      await LocalNotifications.requestPermissions();
    }

    await LocalNotifications.createChannel({
      id: "smart_team_alerts",
      name: "Smart Team Alerts",
      description: "Live Farmer Tracking & Broiler Sale Alerts",
      importance: 5,
      visibility: 1,
      vibration: true,
    });
  } catch (err) {
    console.warn("Permission init warning:", err);
  }
}

// ২. ফোনের নোটিফিকেশন প্যানেলে মেসেজ পাঠানো
export async function triggerPhoneNotification(
  title: string,
  body: string,
  targetRoute: string,
  uniqueId?: number
) {
  if (!Capacitor.isNativePlatform()) return;
  try {
    const perm = await LocalNotifications.checkPermissions();
    if (perm.display !== "granted") {
      const req = await LocalNotifications.requestPermissions();
      if (req.display !== "granted") return;
    }

    const notifId =
      uniqueId || Math.floor(100000 + Math.random() * 900000);

    await LocalNotifications.schedule({
      notifications: [
        {
          id: notifId,
          title,
          body,
          channelId: "smart_team_alerts",
          smallIcon: "res://ic_stat_notification",
          schedule: { at: new Date(Date.now() + 300) },
          extra: { route: targetRoute },
        },
      ],
    });
  } catch (err) {
    console.error("Failed to schedule notification:", err);
  }
}

// ৩. গ্লোবাল Locate Me ফিক্স (কোনো পেজে হাত না দিয়েই সব পেজের GPS কাজ করবে)
let geoPatched = false;
export function setupGlobalNativeGeolocation() {
  if (!Capacitor.isNativePlatform() || geoPatched) return;
  if (typeof navigator === "undefined" || !navigator.geolocation) return;
  geoPatched = true;

  navigator.geolocation.getCurrentPosition = async (
    successCallback,
    errorCallback,
    options
  ) => {
    try {
      const perm = await Geolocation.checkPermissions();
      if (perm.location !== "granted" && perm.coarseLocation !== "granted") {
        const req = await Geolocation.requestPermissions();
        if (req.location !== "granted" && req.coarseLocation !== "granted") {
          errorCallback?.({
            code: 1,
            message: "Location permission denied",
            PERMISSION_DENIED: 1,
            POSITION_UNAVAILABLE: 2,
            TIMEOUT: 3,
          } as GeolocationPositionError);
          return;
        }
      }

      const pos = await Geolocation.getCurrentPosition({
        enableHighAccuracy: options?.enableHighAccuracy ?? true,
        timeout: options?.timeout ?? 15000,
      });

      successCallback(pos as unknown as GeolocationPosition);
    } catch (err: any) {
      errorCallback?.({
        code: 2,
        message: err?.message || "Position unavailable",
        PERMISSION_DENIED: 1,
        POSITION_UNAVAILABLE: 2,
        TIMEOUT: 3,
      } as GeolocationPositionError);
    }
  };
}

// ArrayBuffer থেকে Base64 কনভার্টার
function arrayBufferToBase64(buffer: ArrayBuffer): string {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode.apply(
      null,
      Array.from(bytes.subarray(i, i + chunkSize))
    );
  }
  return window.btoa(binary);
}

export async function saveBlobToNativeDevice(
  blob: Blob,
  suggestedName?: string
) {
  try {
    const buffer = await blob.arrayBuffer();
    const base64Data = arrayBufferToBase64(buffer);

    let ext = ".xlsx";
    if (blob.type.includes("pdf")) ext = ".pdf";
    else if (blob.type.includes("csv")) ext = ".csv";

    const cleanName = (
      suggestedName && suggestedName.trim()
        ? suggestedName.trim()
        : `Report_${new Date().toISOString().slice(0, 10)}${ext}`
    ).replace(/[\\/:*?"<>|]/g, "_");

    // অ্যান্ড্রয়েডের লেটেস্ট ভার্সনের জন্য Documents ফোল্ডার সবচেয়ে নিরাপদ (কোনো এক্সট্রা পারমিশন লাগে না)
    const res = await Filesystem.writeFile({
      path: cleanName,
      data: base64Data,
      directory: Directory.Documents,
      recursive: true,
    });

    alert(`ফাইল সফলভাবে সেভ হয়েছে:\n📁 Documents/${cleanName}`);

    // ফাইলটি সাথে সাথে ওপেন বা শেয়ার করার ডায়ালগ ওপেন করা
    if (res.uri) {
      try {
        await Share.share({
          title: cleanName,
          url: res.uri,
          dialogTitle: "Open or Share Exported File",
        });
      } catch {
        // ইউজার শেয়ার বক্স কেটে দিলে ইগনোর করবে
      }
    }
  } catch (err: any) {
    console.error("File save error:", err);
    alert("ফাইল সেভ করতে সমস্যা হয়েছে: " + (err?.message || ""));
  }
}

// ৪. গ্লোবাল Download ফিক্স (কোনো পেজে হাত না দিয়েই সব পেজের Excel/PDF ডাউনলোড কাজ করবে)
let downloadPatched = false;
export function setupGlobalNativeFileDownload() {
  if (!Capacitor.isNativePlatform() || downloadPatched) return;
  if (typeof window === "undefined") return;
  downloadPatched = true;

  const blobRegistry = new Map<string, Blob>();

  // URL.createObjectURL ইন্টারসেপ্ট করে Blob স্টোর করে রাখা
  const origCreateObjectURL = URL.createObjectURL.bind(URL);
  URL.createObjectURL = function (obj: Blob | MediaSource): string {
    const url = origCreateObjectURL(obj);
    if (obj instanceof Blob) {
      blobRegistry.set(url, obj);
      setTimeout(() => blobRegistry.delete(url), 120000);
    }
    return url;
  };

  // <a download="..."> ক্লিক ইন্টারসেপ্ট করা (ExcelJS, file-saver, jsPDF সব এখানে ধরা পড়বে)
  const origAnchorClick = HTMLAnchorElement.prototype.click;
  HTMLAnchorElement.prototype.click = function () {
    const href = this.getAttribute("href") || this.href || "";
    const downloadName = this.getAttribute("download") || this.download || "";

    if (
      href.startsWith("blob:") ||
      href.startsWith("data:") ||
      this.hasAttribute("download")
    ) {
      const trackedBlob = blobRegistry.get(href);
      if (trackedBlob) {
        saveBlobToNativeDevice(trackedBlob, downloadName);
        return;
      }

      if (href.startsWith("blob:") || href.startsWith("data:")) {
        fetch(href)
          .then((r) => r.blob())
          .then((b) => saveBlobToNativeDevice(b, downloadName))
          .catch(() => origAnchorClick.apply(this));
        return;
      }
    }

    return origAnchorClick.apply(this);
  };
}

export async function getCurrentLatLng(): Promise<[number, number] | null> {
  try {
    // Capacitor নেটিভ প্ল্যাটফর্ম হলে সরাসরি Geolocation প্লাগইন ব্যবহার করব
    if (Capacitor.isNativePlatform()) {
      let perm = await Geolocation.checkPermissions();
      
      if (perm.location !== "granted") {
        perm = await Geolocation.requestPermissions();
      }

      if (perm.location !== "granted" && perm.coarseLocation !== "granted") {
        alert("দয়া করে ফোনের Settings > Apps থেকে এই অ্যাপের Location Permission টি 'Allow only while using the app' বা 'Always' করে দিন।");
        return null;
      }

      const pos = await Geolocation.getCurrentPosition({
        enableHighAccuracy: true,
        timeout: 20000,
        maximumAge: 0,
      });
      return [pos.coords.latitude, pos.coords.longitude];
    } else {
      // ব্রাউজারের জন্য
      return await new Promise((resolve) => {
        if (!navigator.geolocation) {
          alert("ব্রাউজারে লোকেশন সাপোর্ট করছে না।");
          resolve(null);
          return;
        }
        navigator.geolocation.getCurrentPosition(
          (p) => resolve([p.coords.latitude, p.coords.longitude]),
          (err) => {
            console.error(err);
            alert("লোকেশন পাওয়া যায়নি। জিপিএস (GPS) অন আছে কিনা চেক করুন।");
            resolve(null);
          },
          { enableHighAccuracy: true, timeout: 20000 }
        );
      });
    }
  } catch (err: any) {
    console.error("GPS Error:", err);
    alert("লোকেশন ফেচ করতে সমস্যা হয়েছে: " + (err?.message || ""));
    return null;
  }
}