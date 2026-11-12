import {
  LayersControl,
  TileLayer,
} from "react-leaflet";

export default function MapLayers() {

  return (
    <LayersControl position="topright">

      {/* Normal */}

      <LayersControl.BaseLayer
        checked
        name="🗺 Normal"
      >
        <TileLayer
          url="https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}"
          maxZoom={20}
          attribution="&copy; Google Maps"
        />
      </LayersControl.BaseLayer>

      {/* Satellite */}

      <LayersControl.BaseLayer
        name="🛰 Satellite"
      >
        <TileLayer
          url="https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}"
          maxZoom={20}
          attribution="&copy; Google Maps"
        />
      </LayersControl.BaseLayer>

      {/* Hybrid */}

      <LayersControl.BaseLayer
        name="🌍 Hybrid"
      >
        <TileLayer
          url="https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}"
          maxZoom={20}
          attribution="&copy; Google Maps"
        />
      </LayersControl.BaseLayer>

    </LayersControl>
  );
}