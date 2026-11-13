// src/utils/avatarHelper.ts
import React from "react";

// ১. যেকোনো বড় ছবিকে 80x80px (মাত্র ২-৪ KB) ডাটাতে রূপান্তর করার ফাংশন
export function compressImageToTinyBase64(file: File, maxSize = 80): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        canvas.width = maxSize;
        canvas.height = maxSize;
        const ctx = canvas.getContext("2d");
        if (!ctx) return reject("Canvas error");

        const minDim = Math.min(img.width, img.height);
        const sx = (img.width - minDim) / 2;
        const sy = (img.height - minDim) / 2;

        ctx.drawImage(img, sx, sy, minDim, minDim, 0, 0, maxSize, maxSize);
        resolve(canvas.toDataURL("image/webp", 0.65));
      };
      img.onerror = reject;
      img.src = e.target?.result as string;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ২. মেসেজের ভেতর @[Name] থাকলে সেটিকে সুন্দর ব্যাজ হিসেবে দেখানোর ফাংশন
export function renderNoteWithMentions(text: string): React.ReactNode {
  if (!text) return null;
  const parts = text.split(/(@\[[^\]]+\])/g);
  return parts.map((part, index) => {
    if (part.startsWith("@[") && part.endsWith("]")) {
      const name = part.slice(2, -1);
      return React.createElement(
        "span",
        {
          key: index,
          className:
            "inline-flex items-center px-1.5 py-0.2 mx-0.5 rounded bg-blue-100 text-blue-800 font-bold text-[11px] border border-blue-200",
        },
        `@${name}`
      );
    }
    return React.createElement("span", { key: index }, part);
  });
}

// ৩. ছোট প্রোফাইল আইকন কম্পোনেন্ট (ছবি না থাকলে নামের প্রথম অক্ষর দেখাবে)
export function UserAvatar({
  name,
  avatarUrl,
  size = "w-6 h-6",
  textSize = "text-[10px]",
}: {
  name?: string;
  avatarUrl?: string | null;
  size?: string;
  textSize?: string;
}): React.ReactElement {
  const initial = (name || "U").trim().charAt(0).toUpperCase();
  if (avatarUrl) {
    return React.createElement("img", {
      src: avatarUrl,
      alt: name || "User",
      className: `${size} rounded-full object-cover border border-slate-300 shrink-0 shadow-2xs`,
    });
  }
  return React.createElement(
    "div",
    {
      className: `${size} ${textSize} rounded-full bg-emerald-600 text-white font-bold flex items-center justify-center shrink-0 shadow-2xs`,
    },
    initial
  );
}