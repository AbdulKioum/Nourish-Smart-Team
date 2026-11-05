import "leaflet/dist/leaflet.css";
import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import './index.css'
import App from './App.tsx'
import OfflineSyncManager from "./components/OfflineSyncManager";

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
    <OfflineSyncManager />
  </StrictMode>,
)
