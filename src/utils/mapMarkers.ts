import L from "leaflet";

export function getFarmIcon(farmType: string) {
  switch (farmType?.toLowerCase()) {
    case "layer":
      return "🥚";
    case "broiler":
      return "🐔";
    case "cattle":
      return "🐄";
    case "goat":
      return "🐐";
    case "fish":
      return "🐟";
    default:
      return "📍";
  }
}

// ---------------------------------------------------------
// FARMER MARKER
// ---------------------------------------------------------
export function createFarmMarker(
  logoUrl: string,
  farmType: string,
  farmerName: string
) {
  const icon = getFarmIcon(farmType);

  return L.divIcon({
    html: `
      <div style="display:flex; flex-direction:column; align-items:center;">
        <div style="
            background:white;
            color: #1e293b;
            padding:1px 5px;
            border-radius:5px;
            font-size:9px;
            font-weight:700;
            white-space:nowrap;
            box-shadow: 0 2px 4px rgba(0,0,0,0.25);
            margin-bottom:-4px;
            z-index:999;
            position:relative;
          ">
          ${farmerName}
        </div>

        <div style="display:flex; align-items:flex-end; margin-top:-2px;">
          <img src="${logoUrl}" style="
            width: 45px; 
            height: 29px; 
            object-fit: contain;
            /* এই ফিল্টারের মাধ্যমে শুধু লোগোর চারপাশেই সাদা শ্যাডো/স্ট্রোক পড়বে, কোনো চারকোনা বক্স থাকবে না */
            filter: drop-shadow(0px 0px 2px rgba(255, 255, 255, 1)) drop-shadow(0px 2px 4px rgba(0,0,0,0.3));
          " />
          <div style="margin-left:-8px; margin-bottom:-1px; font-size:16px; z-index:999;">
            ${icon}
          </div>
        </div>
      </div>
    `,
    className: "",
    iconSize: [80, 55],
  });
}

// ---------------------------------------------------------
// AGENT MARKER (Main Agent)
// ---------------------------------------------------------
export function createAgentMarker(logoUrl: string, agentName: string) {
  return L.divIcon({
    html: `
      <div style="display:flex; flex-direction:column; align-items:center;">
        <div style="
            background:white;
            color: #1e293b;
            padding:1px 5px;
            border-radius:5px;
            font-size:9px;
            font-weight:700;
            white-space:nowrap;
            box-shadow: 0 2px 4px rgba(0,0,0,0.25);
            margin-bottom:-2px;
            z-index:999;
            position:relative;
          ">
          ${agentName}
        </div>

        <img src="${logoUrl}" style="
          width: 45px; 
          height: 29px; 
          object-fit: contain;
          margin-top:-2px;
          filter: drop-shadow(0px 0px 2px rgba(255, 255, 255, 1)) drop-shadow(0px 2px 4px rgba(0,0,0,0.3));
        " />
      </div>
    `,
    className: "",
    iconSize: [80, 55],
  });
}

// ---------------------------------------------------------
// SUB-AGENT MARKER
// ---------------------------------------------------------
export function createSubAgentMarker(logoUrl: string, subAgentName: string) {
  return L.divIcon({
    html: `
      <div style="display:flex; flex-direction:column; align-items:center;">
        <div style="
            background:white;
            color: #1e293b;
            padding:1px 5px;
            border-radius:5px;
            font-size:9px;
            font-weight:700;
            white-space:nowrap;
            box-shadow: 0 2px 4px rgba(0,0,0,0.25);
            margin-bottom:-2px;
            z-index:999;
            position:relative;
          ">
          ${subAgentName}
        </div>

        <img src="${logoUrl}" style="
          width: 45px; 
          height: 29px; 
          object-fit: contain;
          margin-top:-2px;
          filter: drop-shadow(0px 0px 2px rgba(255, 255, 255, 1)) drop-shadow(0px 2px 4px rgba(0,0,0,0.3));
        " />
      </div>
    `,
    className: "",
    iconSize: [80, 55],
  });
}

// ---------------------------------------------------------
// RETAILER MARKER
// ---------------------------------------------------------
export function createRetailerMarker(logoUrl: string, retailerName: string) {
  return L.divIcon({
    html: `
      <div style="display:flex; flex-direction:column; align-items:center;">
        <div style="
            background:white;
            color: #1e293b;
            padding:1px 5px;
            border-radius:5px;
            font-size:9px;
            font-weight:700;
            white-space:nowrap;
            box-shadow: 0 2px 4px rgba(0,0,0,0.25);
            margin-bottom:-2px;
            z-index:999;
            position:relative;
          ">
          ${retailerName}
        </div>

        <img src="${logoUrl}" style="
          width: 45px; 
          height: 29px; 
          object-fit: contain;
          margin-top:-2px;
          filter: drop-shadow(0px 0px 2px rgba(255, 255, 255, 1)) drop-shadow(0px 2px 4px rgba(0,0,0,0.3));
        " />
      </div>
    `,
    className: "",
    iconSize: [80, 55],
  });
}