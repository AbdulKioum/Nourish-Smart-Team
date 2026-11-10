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

// ফোনের Download / Documents ফোল্ডারে সেভ করার মূল ফাংশন
export async function saveBlobToNativeDevice(
  blob: Blob,
  suggestedName?: string
) {
  try {
    const buffer = await blob.arrayBuffer();
    const base64Data = arrayBufferToBase64(buffer);

    // ফাইলের নাম নির্ধারণ
    let ext = ".xlsx";
    if (blob.type.includes("pdf")) ext = ".pdf";
    else if (blob.type.includes("csv")) ext = ".csv";
    else if (blob.type.includes("png")) ext = ".png";
    else if (blob.type.includes("jpeg") || blob.type.includes("jpg")) ext = ".jpg";

    const cleanName = (
      suggestedName && suggestedName.trim()
        ? suggestedName.trim()
        : `Report_${new Date().toISOString().slice(0, 10)}${ext}`
    ).replace(/[\\/:*?"<>|]/g, "_");

    let savedUri = "";
    let folderLabel = "Download";

    try {
      // প্রথমে সরাসরি ফোনের Download ফোল্ডারে সেভ করার চেষ্টা করবে
      const res = await Filesystem.writeFile({
        path: `Download/${cleanName}`,
        data: base64Data,
        directory: Directory.ExternalStorage,
        recursive: true,
      });
      savedUri = res.uri;
    } catch {
      // Android 11+ যদি ExternalStorage আটকে দেয়, তবে Documents ফোল্ডারে সেভ করবে
      folderLabel = "Documents";
      const resDoc = await Filesystem.writeFile({
        path: cleanName,
        data: base64Data,
        directory: Directory.Documents,
        recursive: true,
      });
      savedUri = resDoc.uri;
    }

    alert(`ফাইলটি ফোনের ${folderLabel} ফোল্ডারে সেভ হয়েছে:\n${cleanName}`);

    // ফাইলটি সাথে সাথে ওপেন বা শেয়ার করার অপশন
    if (savedUri) {
      try {
        await Share.share({
          title: cleanName,
          url: savedUri,
          dialogTitle: "Open or Share File",
        });
      } catch {
        // ইউজার শেয়ার শীট বন্ধ করলে কিছু করার দরকার নেই
      }
    }
  } catch (err: any) {
    alert("ফাইল ডাউনলোড করতে সমস্যা হয়েছে: " + (err?.message || ""));
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