import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { jsPDF } from "jspdf";
import html2canvas from "html2canvas";
import { MapContainer, TileLayer, CircleMarker, Popup, useMap } from "react-leaflet";
import "leaflet/dist/leaflet.css";
import { invoke } from "./localDb";

// Online Swedish voice control via the browser Web Speech API (works in
// Chrome/Edge). This is the same engine the original app used.
const SpeechRecognitionApi = typeof window !== "undefined"
  ? (window.SpeechRecognition || window.webkitSpeechRecognition)
  : null;

// Human-readable heading for each report view (shared by the on-screen title
// and the PDF export).
// SLU stores vernacular species names in lower case; show them with an upper
// case first letter (matching the dropdown labels) wherever they are displayed.
const capitalizeFirst = (value) => {
  const str = String(value ?? "");
  return str ? str.charAt(0).toUpperCase() + str.slice(1) : str;
};
// Cell value for the result tables: capitalise the species column, pass others
// through unchanged.
const speciesAwareCell = (row, key) => (key === "species" ? capitalizeFirst(row[key]) : row[key]);

const labelForViewMode = (mode) =>
  mode === "graf" ? "Tid per år"
    : mode === "tid" ? "Tid"
    : mode === "fenologi" ? "Fenologi"
    : mode === "landskap" ? "Landskap"
    : mode === "kommun" ? "Kommun"
    : mode === "arter" ? "Arter"
    : mode === "karta" ? "Karta"
    : "Resultat";

function MapViewportController({ targetBounds, viewportToken, mapRefExternal }) {
  const map = useMap();

  useEffect(() => {
    mapRefExternal.current = map;
  }, [map, mapRefExternal]);

  useEffect(() => {
    if (!targetBounds) return;
    map.fitBounds(targetBounds, { padding: [30, 30] });
  }, [map, targetBounds, viewportToken]);

  return null;
}

export default function TrollslandeApp() {
  const specialAllOptions = [
    "-- Alla trollsländor --",
    "-- Alla egentliga trollsländor --",
    "-- Alla jungfrusländor --",
    "-- Alla smaragdflicksländor --",
    "-- Alla flicksländor --"
  ];

  const speciesGroups = {
    "-- Alla egentliga trollsländor --": [
      "Bandad ängstrollslända", "Blodröd ängstrollslända", "Blå kejsartrollslända", "Blågrön mosaikslända",
      "Bred kärrtrollslända", "Bred trollslända", "Brun kejsartrollslända", "Brun mosaikslända",
      "Citronfläckad kärrtrollslända", "Fjällmosaikslända", "Fjälltrollslända", "Fyrfläckad trollslända",
      "Grön flodtrollslända", "Grön mosaikslända", "Guldtrollslända", "Gulfläckad glanstrollslända",
      "Gulfläckad ängstrollslända", "Gungflymosaikslända", "Höstmosaikslända", "Karmintrollslända",
      "Kilfläckslända", "Klarblå mosaikslända", "Kungstrollslända", "Metalltrollslända",
      "Mindre glanstrollslända", "Mindre kejsartrollslända", "Mindre sjötrollslända", "Myrtrollslända",
      "Nordisk kärrtrollslända", "Pudrad kärrtrollslända", "Sandflodtrollslända", "Spetsfläckad trollslända",
      "Starrmosaikslända", "Stenflodtrollslända", "Större sjötrollslända", "Större ängstrollslända",
      "Svart ängstrollslända", "Tegelröd ängstrollslända", "Tidig mosaikslända", "Tundratrollslända",
      "Tvåfläckad trollslända", "Vandrande ängstrollslända", "Vassmosaikslända"
    ],
    "-- Alla jungfrusländor --": [
      "Blå jungfruslända", "Blåbandad jungfruslända"
    ],
    "-- Alla smaragdflicksländor --": [
      "Kraftig smaragdflickslända", "Mindre smaragdflickslända", "Pudrad smaragdflickslända", "Vandrande smaragdflickslända"
    ],
    "-- Alla flicksländor --": [
      "Dvärgflickslända", "Flodflickslända", "Griptångsflickslända", "Ljus lyrflickslända",
      "Mindre kustflickslända", "Mindre rödögonflickslända", "Myrflickslända", "Månflickslända",
      "Mörk lyrflickslända", "Röd flickslända", "Sibirisk vinterflickslända", "Sjöflickslända",
      "Spjutflickslända", "Större kustflickslända", "Större rödögonflickslända", "Vinterflickslända"
    ]
  };

  const allSpecies = [
    ...speciesGroups["-- Alla egentliga trollsländor --"],
    ...speciesGroups["-- Alla jungfrusländor --"],
    ...speciesGroups["-- Alla smaragdflicksländor --"],
    ...speciesGroups["-- Alla flicksländor --"]
  ].sort((a, b) => a.localeCompare(b, "sv"));

  const speciesOptions = [...specialAllOptions, ...allSpecies];

  const landscapeOptions = [
    "Blekinge", "Bohuslän", "Dalarna", "Dalsland", "Gotland", "Gästrikland", "Halland", "Hälsingland",
    "Härjedalen", "Jämtland", "Lappland", "Medelpad", "Norrbotten", "Närke", "Skåne", "Småland",
    "Södermanland", "Uppland", "Värmland", "Västerbotten", "Västergötland", "Västmanland", "Ångermanland",
    "Öland", "Östergötland"
  ].sort((a, b) => a.localeCompare(b, "sv"));

  const municipalityOptions = [
    "Ale", "Alingsås", "Alvesta", "Aneby", "Arboga", "Arjeplog", "Arvidsjaur", "Arvika", "Askersund", "Avesta",
    "Bengtsfors", "Berg", "Bjurholm", "Bjuv", "Boden", "Bollebygd", "Bollnäs", "Borgholm", "Borlänge", "Borås",
    "Botkyrka", "Boxholm", "Bromölla", "Bräcke", "Burlöv", "Båstad", "Dals-Ed", "Danderyd", "Degerfors", "Dorotea",
    "Eda", "Ekerö", "Eksjö", "Emmaboda", "Enköping", "Eskilstuna", "Eslöv", "Essunga", "Fagersta", "Falkenberg",
    "Falköping", "Falun", "Filipstad", "Finspång", "Flen", "Forshaga", "Färgelanda", "Gagnef", "Gislaved", "Gnesta",
    "Gnosjö", "Gotland", "Grums", "Grästorp", "Gullspång", "Gällivare", "Gävle", "Göteborg", "Götene", "Habo",
    "Hagfors", "Hallsberg", "Hallstahammar", "Halmstad", "Hammarö", "Haninge", "Haparanda", "Heby", "Hedemora", "Helsingborg",
    "Herrljunga", "Hjo", "Hofors", "Huddinge", "Hudiksvall", "Hultsfred", "Hylte", "Håbo", "Hällefors", "Härjedalen",
    "Härnösand", "Härryda", "Hässleholm", "Höganäs", "Högsby", "Hörby", "Höör", "Jokkmokk", "Järfälla", "Jönköping",
    "Kalix", "Kalmar", "Karlsborg", "Karlshamn", "Karlskoga", "Karlskrona", "Karlstad", "Katrineholm", "Kil", "Kinda",
    "Kiruna", "Klippan", "Knivsta", "Kramfors", "Kristianstad", "Kristinehamn", "Krokom", "Kumla", "Kungsbacka", "Kungsör",
    "Kungälv", "Kävlinge", "Köping", "Laholm", "Landskrona", "Laxå", "Lekeberg", "Leksand", "Lerum", "Lessebo",
    "Lidingö", "Lidköping", "Lilla Edet", "Lindesberg", "Linköping", "Ljungby", "Ljusdal", "Lomma", "Ludvika", "Luleå",
    "Lund", "Lycksele", "Lysekil", "Malmö", "Malung-Sälen", "Malå", "Mariestad", "Mark", "Markaryd", "Mellerud",
    "Mjölby", "Mora", "Motala", "Mullsjö", "Munkedal", "Munkfors", "Mölndal", "Mönsterås", "Mörbylånga", "Nacka",
    "Nora", "Norberg", "Nordanstig", "Nordmaling", "Norrköping", "Norrtälje", "Norsjö", "Nybro", "Nykvarn", "Nyköping",
    "Nynäshamn", "Nässjö", "Ockelbo", "Olofström", "Orsa", "Orust", "Osby", "Oskarshamn", "Ovanåker", "Oxelösund",
    "Pajala", "Partille", "Perstorp", "Piteå", "Ragunda", "Robertsfors", "Ronneby", "Rättvik", "Sala", "Salem",
    "Sandviken", "Sigtuna", "Simrishamn", "Sjöbo", "Skara", "Skellefteå", "Skinnskatteberg", "Skurup", "Skövde", "Smedjebacken",
    "Sollefteå", "Sollentuna", "Solna", "Sorsele", "Sotenäs", "Staffanstorp", "Stenungsund", "Stockholm", "Storfors", "Storuman",
    "Strängnäs", "Strömstad", "Strömsund", "Sundbyberg", "Sundsvall", "Sunne", "Surahammar", "Svalöv", "Svedala", "Svenljunga",
    "Säffle", "Säter", "Sävsjö", "Söderhamn", "Söderköping", "Södertälje", "Sölvesborg", "Tanum", "Tibro", "Tidaholm",
    "Tierp", "Timrå", "Tingsryd", "Tjörn", "Tomelilla", "Torsby", "Torsås", "Tranemo", "Tranås", "Trelleborg",
    "Trollhättan", "Trosa", "Tyresö", "Täby", "Töreboda", "Uddevalla", "Ulricehamn", "Umeå", "Upplands Väsby", "Upplands-Bro",
    "Uppsala", "Uppvidinge", "Vadstena", "Vaggeryd", "Valdemarsvik", "Vallentuna", "Vansbro", "Vara", "Varberg", "Vaxholm",
    "Vellinge", "Vetlanda", "Vilhelmina", "Vimmerby", "Vindeln", "Vingåker", "Vårgårda", "Vänersborg", "Vännäs", "Värmdö",
    "Värnamo", "Västervik", "Västerås", "Växjö", "Ydre", "Ystad", "Åmål", "Ånge", "Åre", "Årjäng", "Åsele",
    "Åstorp", "Åtvidaberg", "Älmhult", "Älvdalen", "Älvkarleby", "Älvsbyn", "Ängelholm", "Öckerö", "Ödeshög", "Örebro",
    "Örkelljunga", "Örnsköldsvik", "Östersund", "Österåker", "Östhammar", "Östra Göinge", "Överkalix", "Övertorneå"
  ].sort((a, b) => a.localeCompare(b, "sv"));

  const monthOptions = [
    { value: "", label: "Alla månader" },
    { value: "01", label: "Januari" },
    { value: "02", label: "Februari" },
    { value: "03", label: "Mars" },
    { value: "04", label: "April" },
    { value: "05", label: "Maj" },
    { value: "06", label: "Juni" },
    { value: "07", label: "Juli" },
    { value: "08", label: "Augusti" },
    { value: "09", label: "September" },
    { value: "10", label: "Oktober" },
    { value: "11", label: "November" },
    { value: "12", label: "December" }
  ];

  const monthNamesShort = ["Jan", "Feb", "Mar", "Apr", "Maj", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dec"];

  const styles = {
    page: { minHeight: "100vh", background: "#eef2f7", padding: 12, fontFamily: "Arial, Helvetica, sans-serif", color: "#1f2937" },
    shell: { maxWidth: 1700, margin: "0 auto", background: "#f8fafc", borderRadius: 20, padding: 16, boxShadow: "0 10px 30px rgba(0,0,0,0.08)" },
    topBar: { display: "flex", justifyContent: "space-between", gap: 16, alignItems: "center", marginBottom: 16 },
    status: { background: "white", borderRadius: 14, padding: "10px 14px", boxShadow: "0 2px 8px rgba(0,0,0,0.06)", fontSize: 14, minWidth: 320 },
    columns: { display: "grid", gridTemplateColumns: "320px 1fr", gap: 16, alignItems: "start" },
    leftCard: { background: "#dbeafe", borderRadius: 20, padding: 12, boxShadow: "0 2px 10px rgba(0,0,0,0.06)" },
    rightCard: { background: "#dcfce7", borderRadius: 20, padding: 16, boxShadow: "0 2px 10px rgba(0,0,0,0.06)", minWidth: 0 },
    section: { marginBottom: 10, textAlign: "left" },
    label: { display: "block", marginBottom: 4, fontSize: 13, fontWeight: 700 },
    select: { width: "100%", padding: "7px 10px", borderRadius: 12, border: "1px solid #93c5fd", background: "white", fontSize: 14, boxSizing: "border-box" },
    input: { width: "100%", padding: "7px 10px", borderRadius: 12, border: "1px solid #93c5fd", background: "white", fontSize: 14, boxSizing: "border-box" },
    checkboxRow: { display: "flex", alignItems: "center", gap: 8, marginTop: 6, fontSize: 13 },
    smallButton: { padding: "10px 12px", borderRadius: 12, border: "1px solid #cbd5e1", background: "white", cursor: "pointer", fontSize: 14 },
    disabledButton: { background: "#e5e7eb", color: "#9ca3af", borderColor: "#d1d5db", cursor: "not-allowed" },
    tagBox: { minHeight: 40, borderRadius: 12, border: "1px solid #93c5fd", background: "#eff6ff", padding: 8, fontSize: 14, boxSizing: "border-box" },
    tag: { borderRadius: 999, background: "#0f172a", color: "white", padding: "6px 10px", border: "none", cursor: "pointer", fontSize: 12 },
    rowButtonsTop: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6 },
    rowButtonsBottom: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 6, marginTop: 8 },
    rowButtonsThird: { display: "grid", gridTemplateColumns: "1fr 2fr", gap: 6, marginTop: 8 },
    phenologyMarkerCell: { padding: "10px 8px", fontSize: 14, fontWeight: 700, background: "#f0fdf4", borderBottom: "1px solid #dcfce7" },
    primaryButton: { width: "100%", padding: "6px 12px", borderRadius: 12, borderWidth: 1, borderStyle: "solid", borderColor: "#0f172a", background: "#0f172a", color: "white", fontSize: 14, fontWeight: 700, cursor: "pointer" },
    secondaryButton: { width: "100%", padding: "6px 12px", borderRadius: 12, borderWidth: 1, borderStyle: "solid", borderColor: "#0f172a", background: "#0f172a", color: "white", fontSize: 14, fontWeight: 700, cursor: "pointer" },
    headerLine: { display: "grid", gridTemplateColumns: "1fr auto auto", alignItems: "start", gap: 16, paddingBottom: 12, borderBottom: "1px solid #bbf7d0", marginBottom: 12 },
    helperText: { fontSize: 14, color: "#1f2937", textAlign: "center", whiteSpace: "pre-line", lineHeight: 1.25, paddingTop: 4 },
    exportDisabled: { background: "#e5e7eb", color: "#9ca3af", border: "1px solid #d1d5db", cursor: "not-allowed" },
    exportEnabled: { background: "#059669", color: "white", border: "1px solid #059669", cursor: "pointer" },
    tableWrap: { border: "1px solid #86efac", borderRadius: 14, overflow: "hidden", background: "white" },
    scrollArea: { maxHeight: 720, overflowY: "auto", overflowX: "auto" },
    table: { width: "100%", borderCollapse: "collapse", tableLayout: "fixed", minWidth: 1380 },
    th: { position: "sticky", top: 0, background: "#bbf7d0", textAlign: "left", padding: "10px 8px", fontSize: 14, borderBottom: "1px solid #86efac", zIndex: 2, cursor: "pointer" },
    td: { padding: "8px", fontSize: 14, borderBottom: "1px solid #dcfce7", textAlign: "left", verticalAlign: "top", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
    empty: { padding: 24, fontSize: 14, color: "#64748b", textAlign: "left" },
    yearRow: { display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 8, alignItems: "start" },
    chartWrap: { border: "1px solid #86efac", borderRadius: 14, background: "white", padding: 16, overflowX: "auto" },
    chartSvg: { display: "block", width: "100%" },
    mapWrap: { border: "1px solid #86efac", borderRadius: 14, overflow: "hidden", background: "white" },
    voiceBox: { border: "1px solid #93c5fd", background: "#eff6ff", borderRadius: 12, padding: 10, fontSize: 13, marginBottom: 8, minHeight: 64 },
    voiceText: { marginTop: 6, color: "#334155", lineHeight: 1.35 }
  };


  const getInitialYear = () => {
    const now = new Date();
    const year = now.getFullYear();
    const mayFirst = new Date(year, 4, 1);
    return String(now > mayFirst ? year : year - 1);
  };


  const [selectedSpecies, setSelectedSpecies] = useState("");
  const [selectedLandscapes, setSelectedLandscapes] = useState([]);
  const [selectedMunicipalities, setSelectedMunicipalities] = useState([]);
  const [observerFilter, setObserverFilter] = useState("");
  const [fromYear, setFromYear] = useState(() => getInitialYear());
  const [toYear, setToYear] = useState(() => getInitialYear());
  const [month, setMonth] = useState("");
  const excludeLarvae = true;
  const [viewMode, setViewMode] = useState("lista");
  const [loading, setLoading] = useState(false);
  const [resultCount, setResultCount] = useState(0);
  const [results, setResults] = useState([]);
  const [statusText, setStatusText] = useState("Välj art och filter. Klicka sedan på en rapportknapp.");
  const [sortField, setSortField] = useState("date");
  const [sortDirection, setSortDirection] = useState("asc");
  const [voiceEnabled, setVoiceEnabled] = useState(false);
  const [voiceListening, setVoiceListening] = useState(false);
  const [lastHeard, setLastHeard] = useState("-");
  const recognitionRef = useRef(null);
  const shouldRestartRef = useRef(false);
  const applyVoiceCommandRef = useRef(null);
  const filtersRef = useRef(null);
  const currentResultsRef = useRef([]);
  const currentViewModeRef = useRef("lista");
  const currentSortFieldRef = useRef("date");
  const currentSortDirectionRef = useRef("asc");
  const mapRef = useRef(null);
  const reportRef = useRef(null);
  const initialMapBoundsRef = useRef(null);
  const [mapTargetBounds, setMapTargetBounds] = useState(null);
  const [mapViewportToken, setMapViewportToken] = useState(0);
  const [viewportHeight, setViewportHeight] = useState(() => (typeof window !== "undefined" ? window.innerHeight : 1080));

  const [useLocalDb, setUseLocalDb] = useState(false);
  const [localDbInfo, setLocalDbInfo] = useState({ exists: false, createdAt: null, rowCount: 0, fileSize: 0, inProgress: false, progressPct: 0, progressText: "" });

  const normalizeVoiceText = (value) => String(value || "")
    .toLowerCase()
    .replaceAll("å", "a")
    .replaceAll("ä", "a")
    .replaceAll("ö", "o")
    .replace(/[^a-z0-9\s-]/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const levenshtein = (a, b) => {
    const s = normalizeVoiceText(a).replaceAll("-", " ");
    const t = normalizeVoiceText(b).replaceAll("-", " ");
    const m = s.length;
    const n = t.length;
    if (!m) return n;
    if (!n) return m;
    const dp = Array.from({ length: m + 1 }, () => Array(n + 1).fill(0));
    for (let i = 0; i <= m; i += 1) dp[i][0] = i;
    for (let j = 0; j <= n; j += 1) dp[0][j] = j;
    for (let i = 1; i <= m; i += 1) {
      for (let j = 1; j <= n; j += 1) {
        const cost = s[i - 1] === t[j - 1] ? 0 : 1;
        dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
      }
    }
    return dp[m][n];
  };

  const yearOptions = useMemo(() => {
    const years = [];
    for (let y = 2026; y >= 1990; y -= 1) years.push(String(y));
    return years;
  }, []);

  const landscapeDisabled = selectedMunicipalities.length > 0;
  const municipalityDisabled = selectedLandscapes.length > 0;

  const selectedSpeciesList = useMemo(() => {
    if (!selectedSpecies) return [];
    if (selectedSpecies === "-- Alla trollsländor --") return allSpecies;
    if (speciesGroups[selectedSpecies]) return speciesGroups[selectedSpecies];
    return [selectedSpecies];
  }, [selectedSpecies]);

  const getSelectedSpeciesListFor = (speciesValue) => {
    if (!speciesValue) return [];
    if (speciesValue === "-- Alla trollsländor --") return allSpecies;
    if (speciesGroups[speciesValue]) return speciesGroups[speciesValue];
    return [speciesValue];
  };

  const isMultiSpeciesSelection = selectedSpeciesList.length > 1;

  const addLandscape = (value) => {
    if (!value || landscapeDisabled || selectedLandscapes.includes(value)) return;
    setSelectedLandscapes([...selectedLandscapes, value]);
  };

  const addMunicipality = (value) => {
    if (!value || municipalityDisabled || selectedMunicipalities.includes(value)) return;
    setSelectedMunicipalities([...selectedMunicipalities, value]);
  };

  const removeItem = (value, selectedValues, setter) => setter(selectedValues.filter((v) => v !== value));
  const escapeCql = (value) => String(value).replaceAll("'", "''");

  const getDateRange = () => {
    if (!month) return { startDate: `${fromYear}-01-01`, endDate: `${toYear}-12-31` };
    const lastDay = new Date(Number(toYear), Number(month), 0).getDate();
    return { startDate: `${fromYear}-${month}-01`, endDate: `${toYear}-${month}-${String(lastDay).padStart(2, "0")}` };
  };

  const buildCurrentConfig = () => ({
    selectedSpecies,
    speciesList: selectedSpeciesList,
    selectedLandscapes,
    selectedMunicipalities,
    fromYear,
    toYear,
    month,
    excludeLarvae,
    observerFilter
  });

  const mapFeatureToRow = (feature) => {
    const p = feature?.properties || {};
    const quantityRaw = p.organismQuantity ?? p.individualCount ?? p.organismQuantityInt ?? "—";
    return {
      occurrenceId: p.occurrenceId || "",
      species: p.vernacularName || "—",
      date: (p.endDate || p.startDate || "").slice(0, 10),
      province: p.province || "—",
      municipality: p.municipality || "—",
      locality: p.locality || p.verbatimLocality || p.locationRemarks || "—",
      quantity: quantityRaw,
      lifeStage: p.lifeStage || "—",
      activity: p.activity || p.behavior || p.reproductiveCondition || "—",
      recordedBy: p.recordedBy || "—",
      latitude: typeof p.decimalLatitude === "number" ? p.decimalLatitude : null,
      longitude: typeof p.decimalLongitude === "number" ? p.decimalLongitude : null,
      isNeverFoundObservation: Boolean(p.isNeverFoundObservation),
      isNotRediscoveredObservation: Boolean(p.isNotRediscoveredObservation)
    };
  };

  const shouldExcludeRow = (row, config = {}) => {
    const configExcludeLarvae = config.excludeLarvae ?? excludeLarvae;
    const configObserverFilter = config.observerFilter ?? observerFilter;
    if (row.isNeverFoundObservation || row.isNotRediscoveredObservation) return true;
    if (configExcludeLarvae) {
      const ls = String(row.lifeStage || "").toLowerCase();
      if (ls.includes("larv") || ls.includes("nymf")) return true;
    }
    if (String(configObserverFilter).trim()) {
      const rf = String(configObserverFilter).trim().toLowerCase();
      if (!String(row.recordedBy || "").toLowerCase().includes(rf)) return true;
    }
    return false;
  };

  const parseQuantityValue = (value) => {
    if (typeof value === "number" && Number.isFinite(value)) return value;
    const raw = String(value ?? "").trim();
    if (!raw || raw === "—") return Number.NEGATIVE_INFINITY;
    const match = raw.replace(",", ".").match(/-?\d+(?:\.\d+)?/);
    return match ? Number(match[0]) : Number.NEGATIVE_INFINITY;
  };

  const compareValues = (a, b, field) => {
    const av = (a?.[field] ?? "").toString().toLowerCase();
    const bv = (b?.[field] ?? "").toString().toLowerCase();
    if (field === "quantity") {
      const aq = parseQuantityValue(a?.[field]);
      const bq = parseQuantityValue(b?.[field]);
      if (aq !== bq) return aq - bq;
      return av.localeCompare(bv, "sv");
    }
    return av.localeCompare(bv, "sv");
  };

  const sortRows = (rows, field = sortField, direction = sortDirection) => {
    const sorted = [...rows].sort((a, b) => compareValues(a, b, field));
    return direction === "asc" ? sorted : sorted.reverse();
  };

  const defaultSortRows = (rows, speciesValue = selectedSpecies) => {
    if (getSelectedSpeciesListFor(speciesValue).length > 1) {
      return [...rows].sort((a, b) => {
        const artCmp = compareValues(a, b, "species");
        if (artCmp !== 0) return artCmp;
        return compareValues(a, b, "date");
      });
    }
    return sortRows(rows, sortField, sortDirection);
  };

  const filterRowsBySelectedMonth = (rows, config = {}) => {
    const configMonth = config.month ?? month;
    const configFromYear = config.fromYear ?? fromYear;
    const configToYear = config.toYear ?? toYear;
    if (!configMonth) return rows;
    return rows.filter((row) => {
      const date = row.date || "";
      return date.slice(0, 4) >= configFromYear && date.slice(0, 4) <= configToYear && date.slice(5, 7) === configMonth;
    });
  };

  const fetchRowsOnline = async (config = {}) => {
    const selectedSpeciesValue = config.selectedSpecies ?? selectedSpecies;
    const speciesList = config.speciesList ?? getSelectedSpeciesListFor(selectedSpeciesValue);
    const configLandscapes = config.selectedLandscapes ?? selectedLandscapes;
    const configMunicipalities = config.selectedMunicipalities ?? selectedMunicipalities;

    const fetchOneBatch = async (speciesName, year) => {
      const yearStart = `${year}-01-01`;
      const yearEnd = `${year}-12-31`;
      const filters = [
        "datasetName='Artportalen'",
        `vernacularName='${escapeCql(speciesName)}'`,
        `endDate >= '${yearStart}'`,
        `endDate <= '${yearEnd}'`
      ];
      if (configLandscapes.length > 0) filters.push(`(${configLandscapes.map((name) => `province='${escapeCql(name)}'`).join(" OR ")})`);
      if (configMunicipalities.length > 0) filters.push(`(${configMunicipalities.map((name) => `municipality='${escapeCql(name)}'`).join(" OR ")})`);

      const params = new URLSearchParams({
        service: "WFS",
        version: "1.0.0",
        request: "GetFeature",
        typeName: "SOS:SpeciesObservations",
        outputFormat: "application/json",
        maxFeatures: "5000",
        CQL_Filter: filters.join(" AND ")
      });
      const response = await fetch(`https://sosgeo.artdata.slu.se/geoserver/SOS/ows?${params.toString()}`);
      if (!response.ok) throw new Error(`WFS-fel ${response.status}`);
      const data = await response.json();
      return Array.isArray(data?.features) ? data.features : [];
    };

    const startYear = Number(config.fromYear ?? fromYear);
    const endYear = Number(config.toYear ?? toYear);
    const years = [];
    for (let year = startYear; year <= endYear; year += 1) years.push(year);

    const featureBatches = [];
    for (const speciesName of speciesList) {
      for (const year of years) {
        const batch = await fetchOneBatch(speciesName, year);
        featureBatches.push(...batch);
      }
    }

    const deduped = [];
    const seen = new Set();
    for (const feature of featureBatches) {
      const occurrenceId = feature?.properties?.occurrenceId || "";
      const dedupeKey = occurrenceId || JSON.stringify(feature?.properties || {});
      if (seen.has(dedupeKey)) continue;
      seen.add(dedupeKey);
      deduped.push(feature);
    }

    return filterRowsBySelectedMonth(deduped.map(mapFeatureToRow), config).filter((row) => !shouldExcludeRow(row, config));
  };

  const normalizeLocalCompare = (value) => String(value || "").normalize("NFC").replace(/\s+/g, " ").trim().toLowerCase();

  const fetchRowsLocal = async (config = {}) => {
    const rows = await invoke("query_local_rows", {
      filters: {
        speciesList: config.speciesList ?? getSelectedSpeciesListFor(config.selectedSpecies ?? selectedSpecies),
        selectedLandscapes: config.selectedLandscapes ?? selectedLandscapes,
        selectedMunicipalities: config.selectedMunicipalities ?? selectedMunicipalities,
        fromYear: config.fromYear ?? fromYear,
        toYear: config.toYear ?? toYear,
        month: config.month ?? month,
        observerFilter: config.observerFilter ?? observerFilter
      }
    });
    const arr = Array.isArray(rows) ? rows : [];
    const speciesSet = new Set((config.speciesList ?? getSelectedSpeciesListFor(config.selectedSpecies ?? selectedSpecies)).map(normalizeLocalCompare));
    const landscapeSet = new Set((config.selectedLandscapes ?? selectedLandscapes).map(normalizeLocalCompare));
    const municipalitySet = new Set((config.selectedMunicipalities ?? selectedMunicipalities).map(normalizeLocalCompare));
    const configFromYear = String(config.fromYear ?? fromYear);
    const configToYear = String(config.toYear ?? toYear);
    const configMonth = String((config.month ?? month) || "");
    const configObserverFilter = normalizeLocalCompare(config.observerFilter ?? observerFilter);

    return arr.filter((row) => {
      const speciesOk = speciesSet.size === 0 || speciesSet.has(normalizeLocalCompare(row.species));
      if (!speciesOk) return false;
      const landscapeOk = landscapeSet.size === 0 || landscapeSet.has(normalizeLocalCompare(row.province));
      if (!landscapeOk) return false;
      const municipalityOk = municipalitySet.size === 0 || municipalitySet.has(normalizeLocalCompare(row.municipality));
      if (!municipalityOk) return false;
      const date = String(row.date || "").trim();
      const year = date.slice(0, 4);
      if (configFromYear && year && year < configFromYear) return false;
      if (configToYear && year && year > configToYear) return false;
      if (configMonth && date.slice(5, 7) !== configMonth) return false;
      if (configObserverFilter && !normalizeLocalCompare(row.recordedBy).includes(configObserverFilter)) return false;
      return !shouldExcludeRow(row, config);
    });
  };

  const getDataRows = async (config = {}) => {
    if (useLocalDb) {
      if (!localDbInfo.exists) throw new Error("Ingen lokal databas finns ännu. Skapa den först.");
      return await fetchRowsLocal(config);
    }
    return await fetchRowsOnline(config);
  };

  const runView = async (mode, label, config = null) => {
    const activeConfig = config || {};
    const selectedSpeciesValue = activeConfig.selectedSpecies ?? selectedSpecies;
    if (!selectedSpeciesValue) {
      setStatusText("Välj först art.");
      return;
    }
    setLoading(true);
    setViewMode(mode);
    setResults([]);
    setResultCount(0);
    try {
      const rows = await getDataRows(activeConfig);
      const finalRows = mode === "lista" ? defaultSortRows(rows, selectedSpeciesValue) : rows;
      setResults(finalRows);
      setResultCount(finalRows.length);
      setStatusText(`Klar. ${label} skapad från ${finalRows.length} observationer för ${selectedSpeciesValue}${useLocalDb ? ' (lokal databas)' : ''}.`);
    } catch (error) {
      setStatusText(`Kunde inte hämta data: ${error.message}`);
      setResults([]);
      setResultCount(0);
    } finally {
      setLoading(false);
    }
  };

  const refreshLocalDbInfo = async () => {
    try {
      const info = await invoke("get_local_db_status");
      setLocalDbInfo(info);
    } catch {
      // ignore
    }
  };

  // Track the most recent `modified` timestamp across fetched features, used as
  // the baseline for later incremental updates.
  const maxModifiedOf = (features, current) => {
    let max = current || "";
    for (const f of (features || [])) {
      const m = f?.properties?.modified;
      if (m && (!max || Date.parse(m) > Date.parse(max))) max = m;
    }
    return max;
  };

  const createLocalDb = async () => {
    const confirmed = window.confirm(
      "Detta RENSAR den befintliga lokala databasen och laddar om ALLA observationer från SLU.\n\n" +
      "Det kan ta lång tid (flera minuter) beroende på antal arter och din uppkoppling.\n\n" +
      "Vill du fortsätta?"
    );
    if (!confirmed) return;

    const currentYear = new Date().getFullYear();
    const FIRST_YEAR = 1990;
    const PAGE = 5000;       // server hard cap of features per request
    const CONCURRENCY = 5;   // global max simultaneous requests (gentle on the API)
    const total = allSpecies.length;
    let done = 0;
    let maxModified = "";

    const updateProgress = async (pct, text) => {
      setLocalDbInfo((prev) => ({ ...prev, inProgress: true, progressPct: pct, progressText: text }));
      try { await invoke("update_local_db_progress", { pct, text }); } catch { /* ignore */ }
    };

    // Global concurrency limiter so the adaptive recursion below never exceeds
    // CONCURRENCY in-flight requests, regardless of how it fans out.
    let active = 0;
    const waiters = [];
    const acquire = () => new Promise((res) => {
      if (active < CONCURRENCY) { active += 1; res(); } else waiters.push(res);
    });
    const release = () => {
      active -= 1;
      if (waiters.length) { active += 1; (waiters.shift())(); }
    };
    const sleep = (ms) => new Promise((res) => setTimeout(res, ms));

    // One page of results, with retry/backoff for transient errors / rate limits.
    const fetchPage = async (cqlFilter, startIndex = 0) => {
      const params = new URLSearchParams({
        service: "WFS",
        version: "2.0.0",
        request: "GetFeature",
        typeNames: "SOS:SpeciesObservations",
        outputFormat: "application/json",
        count: String(PAGE),
        startIndex: String(startIndex),
        CQL_Filter: cqlFilter
      });
      const url = `https://sosgeo.artdata.slu.se/geoserver/SOS/ows?${params.toString()}`;
      let lastErr = null;
      for (let attempt = 0; attempt < 4; attempt += 1) {
        await acquire();
        try {
          const response = await fetch(url);
          if (response.status === 429 || response.status === 503) throw new Error(`Tillfälligt stopp (${response.status})`);
          if (!response.ok) throw new Error(`WFS-fel ${response.status}`);
          const data = await response.json();
          return Array.isArray(data?.features) ? data.features : [];
        } catch (err) {
          lastErr = err;
        } finally {
          release();
        }
        await sleep(1000 * (attempt + 1)); // back off before retrying
      }
      throw lastErr || new Error("Okänt nätverksfel");
    };

    const speciesCql = (species, startY, endY) =>
      `datasetName='Artportalen' AND vernacularName='${escapeCql(species)}' AND endDate >= '${startY}-01-01' AND endDate <= '${endY}-12-31'`;

    // Page through a single year that by itself exceeds the 5000 cap (rare).
    const fetchFullYear = async (species, year) => {
      const all = [];
      let start = 0;
      for (;;) {
        const page = await fetchPage(speciesCql(species, year, year), start);
        all.push(...page);
        if (page.length < PAGE) break;
        start += PAGE;
      }
      return all;
    };

    // Adaptive: fetch the whole year range in one request; only split (and recurse)
    // when a request hits the 5000 cap, i.e. for the few very common species.
    const fetchSpeciesRange = async (species, startY, endY) => {
      const feats = await fetchPage(speciesCql(species, startY, endY), 0);
      if (feats.length < PAGE) return feats;
      if (startY >= endY) return fetchFullYear(species, startY);
      const mid = Math.floor((startY + endY) / 2);
      const [a, b] = await Promise.all([
        fetchSpeciesRange(species, startY, mid),
        fetchSpeciesRange(species, mid + 1, endY)
      ]);
      return a.concat(b);
    };

    try {
      await invoke("prepare_local_db");
      await updateProgress(1, "Förbereder lokal databas...");
      setStatusText("Skapar lokal databas...");

      await Promise.all(allSpecies.map(async (species) => {
        try {
          const features = await fetchSpeciesRange(species, FIRST_YEAR, currentYear);
          maxModified = maxModifiedOf(features, maxModified);
          const rows = features.map(mapFeatureToRow);
          for (let i = 0; i < rows.length; i += 1000) {
            const chunk = rows.slice(i, i + 1000);
            if (chunk.length) await invoke("insert_local_rows_batch", { rows: chunk });
          }
        } finally {
          done += 1;
          const pct = Math.max(1, Math.min(99, (done / total) * 100));
          await updateProgress(pct, `Hämtar arter: ${done} av ${total} klara`);
        }
      }));

      await updateProgress(99, "Slutför lokal databas...");
      const info = await invoke("finalize_local_db", { lastModifiedSync: maxModified });
      setLocalDbInfo(info);
      setStatusText(`Lokal databas klar. ${info.rowCount || 0} poster tillgängliga.`);
    } catch (error) {
      setLocalDbInfo((prev) => ({ ...prev, inProgress: false, progressText: `Fel: ${String(error)}` }));
      setStatusText(`Kunde inte skapa lokal databas: ${String(error)}`);
    }
  };

  // Incremental update: fetch only observations changed since the last sync
  // (by `modified`), which captures both new and after-the-fact edited records.
  const updateLocalDb = async () => {
    try {
      const sync = await invoke("get_sync_info");
      if (!sync?.exists) {
        setStatusText("Ingen lokal databas finns. Klicka på Skapa lokal databas först.");
        return;
      }
      if (!sync.lastModifiedSync) {
        setStatusText("Databasen saknar synk-tidpunkt. Gör en full ombyggnad med Skapa lokal databas en gång.");
        return;
      }
      const since = sync.lastModifiedSync;
      const updateProgress = async (pct, text) => {
        setLocalDbInfo((prev) => ({ ...prev, inProgress: true, progressPct: pct, progressText: text }));
        try { await invoke("update_local_db_progress", { pct, text }); } catch { /* ignore */ }
      };
      const fetchUpdatedBatch = async (speciesName) => {
        const params = new URLSearchParams({
          service: "WFS",
          version: "1.0.0",
          request: "GetFeature",
          typeName: "SOS:SpeciesObservations",
          outputFormat: "application/json",
          maxFeatures: "5000",
          CQL_Filter: [
            "datasetName='Artportalen'",
            `vernacularName='${escapeCql(speciesName)}'`,
            `modified >= '${since}'`
          ].join(" AND ")
        });
        const response = await fetch(`https://sosgeo.artdata.slu.se/geoserver/SOS/ows?${params.toString()}`);
        if (!response.ok) throw new Error(`WFS-fel ${response.status}`);
        const data = await response.json();
        return Array.isArray(data?.features) ? data.features : [];
      };

      let maxModified = since;
      let changedCount = 0;
      try {
        await invoke("begin_local_update");
        await updateProgress(1, "Söker nya och ändrade observationer...");
        setStatusText("Uppdaterar lokal databas...");
        for (let si = 0; si < allSpecies.length; si += 1) {
          const speciesName = allSpecies[si];
          const features = await fetchUpdatedBatch(speciesName);
          maxModified = maxModifiedOf(features, maxModified);
          const rows = features.map(mapFeatureToRow);
          for (let i = 0; i < rows.length; i += 1000) {
            const chunk = rows.slice(i, i + 1000);
            if (chunk.length) {
              await invoke("upsert_local_rows_batch", { rows: chunk });
              changedCount += chunk.length;
            }
          }
          const pct = Math.max(1, Math.min(99, ((si + 1) / allSpecies.length) * 100));
          await updateProgress(pct, `Kontrollerar art ${si + 1} av ${allSpecies.length}: ${speciesName}`);
        }
        await updateProgress(99, "Slutför uppdatering...");
        const info = await invoke("commit_local_update", { lastModifiedSync: maxModified });
        setLocalDbInfo(info);
        setStatusText(`Databasen uppdaterad. ${changedCount} nya/ändrade poster hämtade. Totalt ${info.rowCount || 0} poster.`);
      } catch (error) {
        try { await invoke("rollback_local_update"); } catch { /* ignore */ }
        throw error;
      }
    } catch (error) {
      setLocalDbInfo((prev) => ({ ...prev, inProgress: false, progressText: `Fel: ${String(error)}` }));
      setStatusText(`Kunde inte uppdatera databasen: ${String(error)}`);
    }
  };

  const findBestVoiceOption = (spokenValue, options) => {
    const normalizedSpoken = normalizeVoiceText(spokenValue);
    if (!normalizedSpoken) return null;

    const exact = options.find((option) => normalizeVoiceText(option) === normalizedSpoken);
    if (exact) return exact;

    const contains = options.filter((option) => {
      const normalizedOption = normalizeVoiceText(option);
      return normalizedSpoken.includes(normalizedOption) || normalizedOption.includes(normalizedSpoken);
    });
    if (contains.length > 0) {
      return contains.sort((a, b) => normalizeVoiceText(b).length - normalizeVoiceText(a).length)[0];
    }

    const scored = options.map((option) => {
      const distance = levenshtein(normalizedSpoken, option);
      const maxLen = Math.max(normalizedSpoken.length, normalizeVoiceText(option).length) || 1;
      return { option, score: distance / maxLen };
    }).sort((a, b) => a.score - b.score);

    return scored[0]?.score <= 0.45 ? scored[0].option : null;
  };

  const applyVoiceCommand = async (transcript) => {
    const cleanedTranscript = String(transcript || "").replace(/[,:;]+/g, " ").replace(/\s+/g, " ").trim();
    const state = filtersRef.current || {
      selectedSpecies,
      selectedLandscapes,
      selectedMunicipalities,
      fromYear,
      toYear,
      month,
      excludeLarvae,
      observerFilter
    };

    let speciesValue = state.selectedSpecies;
    const nextLandscapes = [...state.selectedLandscapes];
    const nextMunicipalities = [...state.selectedMunicipalities];
    let nextFromYear = state.fromYear;
    let nextToYear = state.toYear;
    let nextMonth = state.month;
    const changed = [];
    let reportCommand = null;
    let exportRequested = false;

    const normalizedTranscript = normalizeVoiceText(cleanedTranscript);

    if (normalizedTranscript.includes("exportera")) {
      exportRequested = true;
      changed.push("Exportera");
    }
    if (normalizedTranscript.includes("ta bort kommun")) {
      nextMunicipalities.length = 0;
      changed.push("Ta bort kommun");
    }
    if (normalizedTranscript.includes("ta bort landskap")) {
      nextLandscapes.length = 0;
      changed.push("Ta bort landskap");
    }
    if (normalizedTranscript.includes("ta bort manad")) {
      nextMonth = "";
      changed.push("Ta bort månad");
    }

    const sortLabels = {
      "art": "species",
      "datum": "date",
      "landskap": "province",
      "kommun": "municipality",
      "fyndplats": "locality",
      "antal": "quantity",
      "alder stadium": "lifeStage",
      "aktivitet": "activity",
      "observator": "recordedBy"
    };

    const sortMatch = normalizedTranscript.match(/sortera\s+(.+)/);
    if (sortMatch) {
      const spokenSort = sortMatch[1].trim();
      const bestSortLabel = findBestVoiceOption(spokenSort, Object.keys(sortLabels));
      if (bestSortLabel) {
        const sortFieldName = sortLabels[bestSortLabel];
        const nextDirection = currentSortFieldRef.current === sortFieldName && currentSortDirectionRef.current === "asc" ? "desc" : "asc";
        currentSortFieldRef.current = sortFieldName;
        currentSortDirectionRef.current = nextDirection;
        setSortField(sortFieldName);
        setSortDirection(nextDirection);
        if (currentViewModeRef.current === "lista" && currentResultsRef.current.length > 0) {
          setResults(sortRows(currentResultsRef.current, sortFieldName, nextDirection));
          changed.push(`Sortera ${bestSortLabel} ${nextDirection === "asc" ? "stigande" : "fallande"}`);
        } else {
          setStatusText("Visa först en lista innan du sorterar.");
          return;
        }
        return;
      }
    }

    const zoomWords = ["zooma", "summa", "sommar", "dimma"];
    const zoomPrefix = zoomWords.find((word) => normalizedTranscript.startsWith(word));
    if (zoomPrefix) {
      if (currentViewModeRef.current !== "karta") {
        setStatusText("Visa först kartan innan du zoomar.");
        return;
      }

      const zoomArg = normalizedTranscript.slice(zoomPrefix.length).trim();

      if (zoomArg === "ut") {
        if (initialMapBoundsRef.current) {
          setMapTargetBounds(initialMapBoundsRef.current);
          setMapViewportToken((prev) => prev + 1);
          setStatusText("Kartan zoomad ut till sökningens fyndområde.");
        }
        return;
      }

      const normalizedZoomArg = normalizeVoiceText(zoomArg);
      const exactLandscapeMatch = landscapeOptions.find((option) => normalizeVoiceText(option) === normalizedZoomArg);
      if (exactLandscapeMatch) {
        const rows = currentResultsRef.current.filter((row) => row.province === exactLandscapeMatch);
        if (zoomMapToRows(rows, `Kartan zoomad till ${exactLandscapeMatch}.`)) return;
      }

      const exactMunicipalityMatch = municipalityOptions.find((option) => normalizeVoiceText(option) === normalizedZoomArg);
      if (exactMunicipalityMatch) {
        const rows = currentResultsRef.current.filter((row) => row.municipality === exactMunicipalityMatch);
        if (zoomMapToRows(rows, `Kartan zoomad till ${exactMunicipalityMatch}.`)) return;
      }

      const municipalityMatch = findBestVoiceOption(zoomArg, municipalityOptions);
      if (municipalityMatch) {
        const rows = currentResultsRef.current.filter((row) => row.municipality === municipalityMatch);
        if (zoomMapToRows(rows, `Kartan zoomad till ${municipalityMatch}.`)) return;
      }

      const landscapeMatch = findBestVoiceOption(zoomArg, landscapeOptions);
      if (landscapeMatch) {
        const rows = currentResultsRef.current.filter((row) => row.province === landscapeMatch);
        if (zoomMapToRows(rows, `Kartan zoomad till ${landscapeMatch}.`)) return;
      }

      setStatusText("Kunde inte hitta något landskap eller någon kommun att zooma till.");
      return;
    }

    const regex = /(art|landskap|kommun|tid\s+från|tid\s+fran|till|månad|manad|rapport)\s+(.+?)(?=\s+(?:art|landskap|kommun|tid\s+från|tid\s+fran|till|månad|manad|rapport)\b|$)/gi;
    let match;

    while ((match = regex.exec(cleanedTranscript)) !== null) {
      const label = normalizeVoiceText(match[1]);
      const value = match[2].trim();

      if (label === "art") {
        const speciesMatch = findBestVoiceOption(value, speciesOptions);
        if (speciesMatch) {
          speciesValue = speciesMatch;
          changed.push(`Art = ${speciesMatch}`);
        }
        continue;
      }

      if (label === "landskap") {
        const landscapeMatch = findBestVoiceOption(value, landscapeOptions);
        if (landscapeMatch) {
          nextMunicipalities.length = 0;
          if (!nextLandscapes.includes(landscapeMatch)) nextLandscapes.push(landscapeMatch);
          changed.push(`Landskap + ${landscapeMatch}`);
        }
        continue;
      }

      if (label === "kommun") {
        const municipalityMatch = findBestVoiceOption(value, municipalityOptions);
        if (municipalityMatch) {
          nextLandscapes.length = 0;
          if (!nextMunicipalities.includes(municipalityMatch)) nextMunicipalities.push(municipalityMatch);
          changed.push(`Kommun + ${municipalityMatch}`);
        }
        continue;
      }

      if (label === "tid fran") {
        const yearMatch = value.match(/(19|20)\d{2}/);
        if (yearMatch) {
          nextFromYear = yearMatch[0];
          changed.push(`Från år = ${yearMatch[0]}`);
        }
        continue;
      }

      if (label === "till") {
        const yearMatch = value.match(/(19|20)\d{2}/);
        if (yearMatch) {
          nextToYear = yearMatch[0];
          changed.push(`Till år = ${yearMatch[0]}`);
        }
        continue;
      }

      if (label === "manad") {
        const monthMatch = findBestVoiceOption(value, monthOptions.slice(1).map((item) => item.label));
        if (monthMatch) {
          const monthEntry = monthOptions.find((item) => item.label === monthMatch);
          if (monthEntry) {
            nextMonth = monthEntry.value;
            changed.push(`Månad = ${monthEntry.label}`);
          }
        }
        continue;
      }

      if (label === "rapport") {
        const normalizedValue = normalizeVoiceText(value);
        if (normalizedValue.includes("lista")) reportCommand = { mode: "lista", label: "Lista" };
        else if (normalizedValue.includes("tid per ar") || normalizedValue.includes("tid per år") || normalizedValue.includes("tid perar")) reportCommand = { mode: "graf", label: "Tid per år" };
        else if (normalizedValue.includes("tid") || normalizedValue.includes("tidskrav") || normalizedValue.includes("tidsgraf")) reportCommand = { mode: "tid", label: "Tid" };
        else if (normalizedValue.includes("fenologi")) reportCommand = { mode: "fenologi", label: "Fenologi" };
        else if (normalizedValue.includes("kommun")) reportCommand = { mode: "kommun", label: "Kommun" };
        else if (normalizedValue.includes("landskap")) reportCommand = { mode: "landskap", label: "Landskap" };
        else if (normalizedValue.includes("karta")) reportCommand = { mode: "karta", label: "Karta" };
        else if (normalizedValue.includes("arter")) reportCommand = { mode: "arter", label: "Arter" };
        if (reportCommand) changed.push(`Rapport = ${reportCommand.label}`);
      }
    }

    setSelectedSpecies(speciesValue);
    setSelectedLandscapes(nextLandscapes);
    setSelectedMunicipalities(nextMunicipalities);
    setFromYear(nextFromYear);
    setToYear(nextToYear);
    setMonth(nextMonth);

    const nextConfig = {
      selectedSpecies: speciesValue,
      selectedLandscapes: nextLandscapes,
      selectedMunicipalities: nextMunicipalities,
      fromYear: nextFromYear,
      toYear: nextToYear,
      month: nextMonth,
      excludeLarvae,
      observerFilter
    };
    filtersRef.current = nextConfig;

    if (reportCommand) {
      await runView(reportCommand.mode, reportCommand.label, nextConfig);
      return;
    }

    if (exportRequested) {
      if (currentViewModeRef.current === "lista") {
        if (currentResultsRef.current.length > 0) {
          exportToExcel(currentResultsRef.current, currentViewModeRef.current);
        } else {
          setStatusText("Visa först en lista innan du säger Exportera till Excel.");
        }
      } else {
        await exportToPdf();
      }
      return;
    }

    if (changed.length > 0) {
      setStatusText(`Röstkommando förstått: ${changed.join(", ")}`);
    }
  };

  const handleSort = (field) => {
    const newDirection = sortField === field && sortDirection === "asc" ? "desc" : "asc";
    setSortField(field);
    setSortDirection(newDirection);
    setResults((prev) => sortRows(prev, field, newDirection));
  };

  const exportToExcel = (rowsArg = results, modeArg = viewMode) => {
    if (!Array.isArray(rowsArg) || !rowsArg.length || modeArg !== "lista") return;
    const includeSpeciesColumn = new Set(rowsArg.map((row) => row.species).filter(Boolean)).size > 1;
    const header = includeSpeciesColumn
      ? ["Art", "Datum", "Landskap", "Kommun", "Fyndplats", "Antal", "Ålder/stadium", "Aktivitet", "Observatör"]
      : ["Datum", "Landskap", "Kommun", "Fyndplats", "Antal", "Ålder/stadium", "Aktivitet", "Observatör"];
    const body = rowsArg.map((row) => includeSpeciesColumn
      ? [capitalizeFirst(row.species), row.date, row.province, row.municipality, row.locality, row.quantity, row.lifeStage, row.activity, row.recordedBy]
      : [row.date, row.province, row.municipality, row.locality, row.quantity, row.lifeStage, row.activity, row.recordedBy]);
    const worksheet = XLSX.utils.aoa_to_sheet([header, ...body]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Observationer");
    const filename = `trollsländor_${selectedSpecies.replaceAll(" ", "_").replaceAll("--", "")}_${fromYear}${month ? `_${month}` : ""}.xlsx`;
    XLSX.writeFileXLSX(workbook, filename);
    setStatusText("Rapport sparad som riktig Excel-fil (.xlsx).");
  };

  // Export any non-list report (graf/tid/landskap/kommun/karta/fenologi) as a
  // PDF by rasterising the rendered report area. Triggered by the export button
  // and by the voice command "Exportera som PDF".
  const exportToPdf = async () => {
    const node = reportRef.current;
    if (!node || currentResultsRef.current.length === 0) {
      setStatusText("Visa först en rapport innan du exporterar som PDF.");
      return;
    }
    try {
      setStatusText("Skapar PDF...");
      // The map shows the current pan/zoom viewport; give freshly requested
      // tiles a moment to finish loading so the capture matches the screen.
      if (currentViewModeRef.current === "karta") {
        await new Promise((resolve) => setTimeout(resolve, 700));
      }
      const canvas = await html2canvas(node, { backgroundColor: "#ffffff", scale: 2, useCORS: true });
      const imgData = canvas.toDataURL("image/png");
      const orientation = canvas.width >= canvas.height ? "landscape" : "portrait";
      const pdf = new jsPDF({ orientation, unit: "pt", format: "a4" });
      const pageWidth = pdf.internal.pageSize.getWidth();
      const pageHeight = pdf.internal.pageSize.getHeight();
      const margin = 24;
      const monthLabel = month ? (monthOptions.find((m) => m.value === month)?.label || "") : "";
      const period = `${fromYear}${toYear !== fromYear ? `–${toYear}` : ""}${monthLabel ? `, ${monthLabel}` : ""}`;
      const title = `${labelForViewMode(currentViewModeRef.current)} – ${selectedSpecies} (${period})`;
      pdf.setFontSize(13);
      pdf.text(title, margin, margin + 6);
      const headerOffset = margin + 18;
      const availW = pageWidth - margin * 2;
      const availH = pageHeight - headerOffset - margin;
      const ratio = Math.min(availW / canvas.width, availH / canvas.height);
      const w = canvas.width * ratio;
      const h = canvas.height * ratio;
      pdf.addImage(imgData, "PNG", margin, headerOffset, w, h);
      const filename = `trollsländor_${currentViewModeRef.current}_${selectedSpecies.replaceAll(" ", "_").replaceAll("--", "")}_${fromYear}${month ? `_${month}` : ""}.pdf`;
      pdf.save(filename);
      setStatusText("Rapport sparad som PDF.");
    } catch (error) {
      setStatusText(`Kunde inte skapa PDF: ${error?.message || error}`);
    }
  };

  useEffect(() => {
    if (typeof window === "undefined") return undefined;
    const onResize = () => setViewportHeight(window.innerHeight);
    window.addEventListener("resize", onResize);
    return () => window.removeEventListener("resize", onResize);
  }, []);

  useEffect(() => {
    refreshLocalDbInfo();
    const timer = window.setInterval(() => { refreshLocalDbInfo(); }, 500);
    return () => window.clearInterval(timer);
  }, []);

  useEffect(() => {
    if (localDbInfo.inProgress) return;
    if (localDbInfo.exists && localDbInfo.progressText === "Lokal databas klar") {
      setStatusText(`Lokal databas klar. ${localDbInfo.rowCount || 0} poster tillgängliga.`);
    }
  }, [localDbInfo.inProgress, localDbInfo.exists, localDbInfo.progressText, localDbInfo.rowCount]);

  useEffect(() => {
    filtersRef.current = {
      selectedSpecies,
      selectedLandscapes,
      selectedMunicipalities,
      fromYear,
      toYear,
      month,
      excludeLarvae,
      observerFilter
    };
  }, [selectedSpecies, selectedLandscapes, selectedMunicipalities, fromYear, toYear, month, excludeLarvae, observerFilter]);

  useEffect(() => {
    currentResultsRef.current = results;
    currentViewModeRef.current = viewMode;
    currentSortFieldRef.current = sortField;
    currentSortDirectionRef.current = sortDirection;
  }, [results, viewMode, sortField, sortDirection]);

  useEffect(() => {
    if (viewMode !== "karta") return;
    const bounds = rowsToBounds(results);
    initialMapBoundsRef.current = bounds;
    if (bounds) {
      setMapTargetBounds(bounds);
      setMapViewportToken((prev) => prev + 1);
    }
  }, [viewMode, results]);

  // Keep a stable ref to the latest applyVoiceCommand so the long-lived
  // recognizer callback always sees current filter state.
  useEffect(() => {
    applyVoiceCommandRef.current = applyVoiceCommand;
  });

  // Create the SpeechRecognition instance once. onresult/onend read the latest
  // applyVoiceCommand via a ref, so the long-lived instance never goes stale.
  useEffect(() => {
    if (!SpeechRecognitionApi) return undefined;

    const recognition = new SpeechRecognitionApi();
    recognition.lang = "sv-SE";
    recognition.continuous = false;
    recognition.interimResults = false;

    recognition.onstart = () => setVoiceListening(true);
    recognition.onend = () => {
      setVoiceListening(false);
      if (shouldRestartRef.current) {
        window.setTimeout(() => {
          try { recognition.start(); } catch (_) { /* ignore repeated start */ }
        }, 150);
      }
    };
    recognition.onerror = (event) => {
      setVoiceListening(false);
      if (event?.error === "network") {
        setStatusText("Röstfel: ingen internetuppkoppling för taligenkänningen (online krävs).");
      } else if (event?.error === "not-allowed" || event?.error === "service-not-allowed") {
        setStatusText("Mikrofonåtkomst nekad. Tillåt mikrofonen i webbläsaren.");
      }
    };
    recognition.onresult = async (event) => {
      const transcript = Array.from(event.results)
        .map((result) => result[0]?.transcript || "")
        .join(" ")
        .trim();
      if (!transcript) return;
      setLastHeard(transcript);
      const fn = applyVoiceCommandRef.current;
      if (fn) await fn(transcript);
    };

    recognitionRef.current = recognition;
    return () => {
      shouldRestartRef.current = false;
      try { recognition.stop(); } catch (_) { /* ignore */ }
      recognitionRef.current = null;
    };
  }, []);

  // Start/stop listening when the microphone toggle changes.
  useEffect(() => {
    if (!recognitionRef.current) return;
    shouldRestartRef.current = voiceEnabled;

    if (voiceEnabled) {
      try {
        recognitionRef.current.start();
        setStatusText("Mikrofonläge aktivt. Väntar på röstkommando.");
      } catch (_) { /* ignore if already started */ }
    } else {
      try { recognitionRef.current.stop(); } catch (_) { /* ignore */ }
      setVoiceListening(false);
    }
  }, [voiceEnabled]);

  const getQuarterBoundaries = (year, monthIndex) => {
    const days = new Date(year, monthIndex + 1, 0).getDate();
    if (days === 31) return [[1, 8], [9, 16], [17, 24], [25, 31]];
    if (days === 30) return [[1, 8], [9, 15], [16, 23], [24, 30]];
    if (days === 29) return [[1, 7], [8, 14], [15, 21], [22, 29]];
    return [[1, 7], [8, 14], [15, 21], [22, 28]];
  };

  const chartBlocksByYear = useMemo(() => {
    if (viewMode !== "graf" || results.length === 0) return [];
    const observedMonthsByYear = new Map();
    results.forEach((row) => {
      if (!row.date) return;
      const year = Number(row.date.slice(0, 4));
      const monthIndex = Number(row.date.slice(5, 7)) - 1;
      if (!observedMonthsByYear.has(year)) observedMonthsByYear.set(year, new Set());
      observedMonthsByYear.get(year).add(monthIndex);
    });
    const includedMonths = [];
    Array.from(observedMonthsByYear.keys()).sort((a, b) => a - b).forEach((year) => {
      const months = Array.from(observedMonthsByYear.get(year)).sort((a, b) => a - b);
      const minMonth = months[0];
      const maxMonth = months[months.length - 1];
      for (let monthIndex = minMonth; monthIndex <= maxMonth; monthIndex += 1) includedMonths.push({ year, monthIndex });
    });
    const counts = new Map();
    includedMonths.forEach(({ year, monthIndex }) => {
      getQuarterBoundaries(year, monthIndex).forEach((_, quarterIndex) => counts.set(`${year}-${monthIndex}-${quarterIndex}`, 0));
    });
    results.forEach((row) => {
      if (!row.date) return;
      const year = Number(row.date.slice(0, 4));
      const monthIndex = Number(row.date.slice(5, 7)) - 1;
      const day = Number(row.date.slice(8, 10));
      const quarterIndex = getQuarterBoundaries(year, monthIndex).findIndex(([start, end]) => day >= start && day <= end);
      const key = `${year}-${monthIndex}-${quarterIndex}`;
      if (counts.has(key)) counts.set(key, (counts.get(key) || 0) + 1);
    });
    const blocks = [];
    includedMonths.forEach(({ year, monthIndex }) => {
      getQuarterBoundaries(year, monthIndex).forEach((_, quarterIndex) => {
        blocks.push({ key: `${year}-${monthIndex}-${quarterIndex}`, year, monthIndex, monthLabel: monthNamesShort[monthIndex], count: counts.get(`${year}-${monthIndex}-${quarterIndex}`) || 0 });
      });
    });
    return blocks;
  }, [results, viewMode]);

  const chartBlocksCombined = useMemo(() => {
    if (viewMode !== "tid" || results.length === 0) return [];
    const observedMonths = new Set();
    results.forEach((row) => {
      if (!row.date) return;
      observedMonths.add(Number(row.date.slice(5, 7)) - 1);
    });
    const months = Array.from(observedMonths).sort((a, b) => a - b);
    if (months.length === 0) return [];
    const includedMonths = [];
    for (let monthIndex = months[0]; monthIndex <= months[months.length - 1]; monthIndex += 1) includedMonths.push(monthIndex);
    const counts = new Map();
    includedMonths.forEach((monthIndex) => {
      getQuarterBoundaries(2025, monthIndex).forEach((_, quarterIndex) => counts.set(`${monthIndex}-${quarterIndex}`, 0));
    });
    results.forEach((row) => {
      if (!row.date) return;
      const monthIndex = Number(row.date.slice(5, 7)) - 1;
      const day = Number(row.date.slice(8, 10));
      const quarterIndex = getQuarterBoundaries(2025, monthIndex).findIndex(([start, end]) => day >= start && day <= end);
      const key = `${monthIndex}-${quarterIndex}`;
      if (counts.has(key)) counts.set(key, (counts.get(key) || 0) + 1);
    });
    const blocks = [];
    includedMonths.forEach((monthIndex) => {
      getQuarterBoundaries(2025, monthIndex).forEach((_, quarterIndex) => {
        blocks.push({ key: `${monthIndex}-${quarterIndex}`, monthIndex, monthLabel: monthNamesShort[monthIndex], count: counts.get(`${monthIndex}-${quarterIndex}`) || 0 });
      });
    });
    return blocks;
  }, [results, viewMode]);

  const landscapeSeries = useMemo(() => {
    if (viewMode !== "landskap" || results.length === 0) return [];
    const counts = new Map();
    results.forEach((row) => counts.set(row.province || "—", (counts.get(row.province || "—") || 0) + 1));
    return Array.from(counts.entries()).map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count);
  }, [results, viewMode]);

  const municipalitySeries = useMemo(() => {
    if (viewMode !== "kommun" || results.length === 0) return [];
    const counts = new Map();
    results.forEach((row) => counts.set(row.municipality || "—", (counts.get(row.municipality || "—") || 0) + 1));
    return Array.from(counts.entries()).map(([name, count]) => ({ name, count })).sort((a, b) => b.count - a.count);
  }, [results, viewMode]);

  const speciesSeries = useMemo(() => {
    if (viewMode !== "arter" || results.length === 0) return [];
    const counts = new Map();
    results.forEach((row) => counts.set(row.species || "—", (counts.get(row.species || "—") || 0) + 1));
    return Array.from(counts.entries()).map(([name, count]) => ({ name: capitalizeFirst(name), count })).sort((a, b) => b.count - a.count);
  }, [results, viewMode]);

  const activeChartBlocks = viewMode === "graf" ? chartBlocksByYear : viewMode === "tid" ? chartBlocksCombined : [];

  const yStep = useMemo(() => {
    const max = activeChartBlocks.length ? Math.max(...activeChartBlocks.map((d) => d.count), 0) : 0;
    const steps = [1, 2, 5, 10, 20, 50, 100, 200, 500, 1000];
    return steps.find((step) => max <= step * 5) || 1000;
  }, [activeChartBlocks]);

  const yMax = useMemo(() => {
    const max = activeChartBlocks.length ? Math.max(...activeChartBlocks.map((d) => d.count), 0) : 0;
    if (max === 0) return yStep;
    return (Math.floor(max / yStep) + 1) * yStep;
  }, [activeChartBlocks, yStep]);

  const phenologyResults = useMemo(() => {
    if (viewMode !== "fenologi" || results.length === 0) return [];
    const withMonthDay = results
      .filter((row) => row.date)
      .map((row) => {
        const monthDay = row.date.slice(5, 10);
        return { ...row, __monthDay: monthDay };
      });
    const sorted = [...withMonthDay].sort((a, b) => {
      const cmp = a.__monthDay.localeCompare(b.__monthDay, "sv");
      if (cmp !== 0) return cmp;
      return (a.date || "").localeCompare(b.date || "", "sv");
    });
    const earliest = sorted.slice(0, 3).map((row) => ({ ...row, __phenologySection: "tidigast" }));
    const latest = [...sorted].reverse().slice(0, 3).reverse().map((row) => ({ ...row, __phenologySection: "senast" }));
    return [...earliest, ...latest];
  }, [results, viewMode]);

  const SelectedBox = ({ title, items, onRemove, emptyText }) => (
    <div style={styles.section}>
      <div style={styles.label}>{title}</div>
      <div style={styles.tagBox}>
        {items.length === 0 ? emptyText : <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>{items.map((item) => <button key={item} onClick={() => onRemove(item)} style={styles.tag}>{item} ×</button>)}</div>}
      </div>
    </div>
  );

  const columns = isMultiSpeciesSelection
    ? [
        { key: "species", label: "Art", width: "180px" },
        { key: "date", label: "Datum", width: "80px" },
        { key: "province", label: "Landskap", width: "90px" },
        { key: "municipality", label: "Kommun", width: "90px" },
        { key: "locality", label: "Fyndplats", width: "240px" },
        { key: "quantity", label: "Antal", width: "50px" },
        { key: "lifeStage", label: "Ålder/stadium", width: "90px" },
        { key: "activity", label: "Aktivitet", width: "100px" },
        { key: "recordedBy", label: "Observatör", width: "180px" }
      ]
    : [
        { key: "date", label: "Datum", width: "70px" },
        { key: "province", label: "Landskap", width: "90px" },
        { key: "municipality", label: "Kommun", width: "90px" },
        { key: "locality", label: "Fyndplats", width: "240px" },
        { key: "quantity", label: "Antal", width: "50px" },
        { key: "lifeStage", label: "Ålder/stadium", width: "90px" },
        { key: "activity", label: "Aktivitet", width: "100px" },
        { key: "recordedBy", label: "Observatör", width: "180px" }
      ];

  const monthCount = Math.max(1, activeChartBlocks.length / 4);
  const chartWidth = Math.max(980, monthCount * 130 + 100);
  const chartHeight = 460;
  const chartLeft = 56;
  const chartBottom = 66;
  const chartTop = 26;
  const plotHeight = chartHeight - chartTop - chartBottom;
  const groupWidth = 120;
  const barGap = 4;
  const barWidth = 26;

  const horizontalSeries = viewMode === "landskap" ? landscapeSeries : viewMode === "arter" ? speciesSeries : municipalitySeries;
  const horizontalMax = horizontalSeries.length ? Math.max(...horizontalSeries.map((d) => d.count), 1) : 1;
  const horizontalChartWidth = 980;
  const horizontalChartHeight = Math.max(280, horizontalSeries.length * 32 + 80);
  const horizontalLeft = viewMode === "arter" ? 280 : 180;
  const horizontalTop = 20;
  const horizontalRight = 40;
  const horizontalPlotWidth = horizontalChartWidth - horizontalLeft - horizontalRight;

  const mapPoints = useMemo(() => results.filter((row) => row.latitude != null && row.longitude != null), [results]);

  const rowsToBounds = (rows) => {
    const points = rows.filter((row) => row.latitude != null && row.longitude != null);
    if (points.length === 0) return null;

    const latitudes = points.map((row) => Number(row.latitude));
    const longitudes = points.map((row) => Number(row.longitude));
    let minLat = Math.min(...latitudes);
    let maxLat = Math.max(...latitudes);
    let minLng = Math.min(...longitudes);
    let maxLng = Math.max(...longitudes);

    if (minLat === maxLat) {
      minLat -= 0.05;
      maxLat += 0.05;
    }
    if (minLng === maxLng) {
      minLng -= 0.05;
      maxLng += 0.05;
    }

    return [[minLat, minLng], [maxLat, maxLng]];
  };

  const zoomMapToRows = (rows, statusLabel = "") => {
    const bounds = rowsToBounds(rows);
    if (!bounds) {
      if (statusLabel) setStatusText(statusLabel);
      return false;
    }
    setMapTargetBounds(bounds);
    setMapViewportToken((prev) => prev + 1);
    if (statusLabel) setStatusText(statusLabel);
    return true;
  };

  const mapCenter = mapPoints.length > 0 ? [mapPoints[0].latitude, mapPoints[0].longitude] : [62.0, 15.0];
  const landscapeGraphEnabled = selectedLandscapes.length === 0 && selectedMunicipalities.length === 0;
  const municipalityGraphEnabled = selectedMunicipalities.length === 0;
  const dynamicScrollMaxHeight = Math.max(420, Math.min(820, viewportHeight - 185));
  const dynamicMapHeight = Math.max(460, Math.min(820, viewportHeight - 185));

  const viewModeLabel = labelForViewMode(viewMode);
  const currentReportHasContent =
    viewMode === "lista" ? results.length > 0
      : viewMode === "fenologi" ? phenologyResults.length > 0
      : (viewMode === "graf" || viewMode === "tid") ? activeChartBlocks.length > 0
      : (viewMode === "landskap" || viewMode === "kommun" || viewMode === "arter") ? horizontalSeries.length > 0
      : viewMode === "karta" ? mapPoints.length > 0
      : false;
  const exportIsExcel = viewMode === "lista";
  const exportLabel = exportIsExcel ? "Exportera till Excel" : "Exportera som PDF";
  const exportStyle = currentReportHasContent
    ? { ...styles.smallButton, ...styles.exportEnabled }
    : { ...styles.smallButton, ...styles.exportDisabled };

  return (
    <div style={styles.page}>
      <div style={styles.shell}>
        <div style={styles.topBar}>
          <div style={{ display: "flex", alignItems: "center", gap: 14, minWidth: 0 }}>
            <h1 style={{ margin: 0, fontSize: 28, whiteSpace: "nowrap" }}>Trollsländor - Sök i ArtPortalen</h1>
            {localDbInfo.exists ? <div style={{ fontSize: 13, color: "#065f46", whiteSpace: "nowrap" }}>Lokal databas skapad {localDbInfo.createdAt || ""}</div> : null}
          </div>
          <div style={styles.status}>
            <div>{loading ? "Arbetar..." : statusText}</div>
            {localDbInfo.inProgress ? (
              <div style={{ marginTop: 8 }}>
                <div style={{ fontSize: 12, marginBottom: 4 }}>{localDbInfo.progressText || "Skapar lokal databas..."}</div>
                <div style={{ height: 10, background: "#d1d5db", borderRadius: 999, overflow: "hidden" }}>
                  <div style={{ width: `${localDbInfo.progressPct || 0}%`, height: "100%", background: "#2563eb" }} />
                </div>
              </div>
            ) : null}
          </div>
        </div>

        <div style={styles.columns}>
          <div style={styles.leftCard}>
            <div style={styles.voiceBox}>
              <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
                <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
                  <input
                    type="checkbox"
                    checked={voiceEnabled}
                    onChange={(e) => setVoiceEnabled(e.target.checked)}
                    disabled={!SpeechRecognitionApi}
                  />
                  <span>Mikrofonläge</span>
                </label>
                <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13 }}>
                  <input
                    type="checkbox"
                    checked={useLocalDb}
                    onChange={(e) => setUseLocalDb(e.target.checked)}
                    disabled={!localDbInfo.exists && !useLocalDb}
                  />
                  <span>Lokal databas</span>
                </label>
              </div>
              {!SpeechRecognitionApi ? (
                <div style={styles.voiceText}>Röststyrning stöds inte i denna webbläsare (använd Chrome eller Edge).</div>
              ) : null}
              <div style={{ ...styles.voiceText, marginTop: 4 }}>
                {voiceEnabled && voiceListening ? "🎤 Lyssnar – " : ""}Senast hört: {lastHeard || "-"}
              </div>
            </div>

            <div style={styles.section}>
              <div style={{ display: "grid", gridTemplateColumns: "64px 1fr", gap: 8, alignItems: "center" }}>
                <div style={{ fontSize: 13, fontWeight: 700 }}>Art</div>
                <select value={selectedSpecies} onChange={(e) => setSelectedSpecies(e.target.value)} style={styles.select}>
                  <option value="">Välj art</option>
                  {speciesOptions.map((species) => <option key={species} value={species}>{species}</option>)}
                </select>
              </div>
            </div>

            <div style={styles.section}>
              <div style={{ display: "grid", gridTemplateColumns: "74px 1fr", gap: 8, alignItems: "center", marginBottom: 6 }}>
                <div style={{ fontSize: 13, fontWeight: 700 }}>Landskap</div>
                <select value="" onChange={(e) => addLandscape(e.target.value)} style={styles.select} disabled={landscapeDisabled}>
                  <option value="">Välj landskap</option>
                  {landscapeOptions.map((name) => <option key={name} value={name}>{name}</option>)}
                </select>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "74px 1fr", gap: 8, alignItems: "center" }}>
                <div style={{ fontSize: 13, fontWeight: 700 }}>Kommun</div>
                <select value="" onChange={(e) => addMunicipality(e.target.value)} style={styles.select} disabled={municipalityDisabled}>
                  <option value="">Välj kommun</option>
                  {municipalityOptions.map((name) => <option key={name} value={name}>{name}</option>)}
                </select>
              </div>
            </div>

            <SelectedBox title="Valda landskap" items={selectedLandscapes} onRemove={(item) => removeItem(item, selectedLandscapes, setSelectedLandscapes)} emptyText="Inga landskap valda" />
            <SelectedBox title="Valda kommuner" items={selectedMunicipalities} onRemove={(item) => removeItem(item, selectedMunicipalities, setSelectedMunicipalities)} emptyText="Inga kommuner valda" />

            <div style={styles.section}>
              <div style={styles.label}>Observatör</div>
              <input value={observerFilter} onChange={(e) => setObserverFilter(e.target.value)} placeholder="Skriv hela eller del av namn" style={styles.input} />
            </div>

            <div style={styles.section}>
              <div style={styles.yearRow}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 4 }}>Tid från</div>
                  <select value={fromYear} onChange={(e) => setFromYear(e.target.value)} style={styles.select}>{yearOptions.map((y) => <option key={y} value={y}>{y}</option>)}</select>
                </div>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 4 }}>Till</div>
                  <select value={toYear} onChange={(e) => setToYear(e.target.value)} style={styles.select}>{yearOptions.map((y) => <option key={y} value={y}>{y}</option>)}</select>
                </div>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 4 }}>Månad</div>
                  <select value={month} onChange={(e) => setMonth(e.target.value)} style={styles.select}>{monthOptions.map((m) => <option key={m.value} value={m.value}>{m.label}</option>)}</select>
                </div>
              </div>
            </div>

            <div style={styles.section}>
              <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 4 }}>Rapport</div>
              <div style={styles.rowButtonsTop}>
                <button onClick={() => runView("lista", "Lista", buildCurrentConfig())} style={styles.primaryButton}>Lista</button>
                <button onClick={() => runView("tid", "Tid", buildCurrentConfig())} style={styles.primaryButton}>Tid</button>
                <button onClick={() => runView("graf", "Tid per år", buildCurrentConfig())} disabled={fromYear === toYear} style={fromYear === toYear ? { ...styles.primaryButton, ...styles.disabledButton } : styles.primaryButton}>Tid per år</button>
              </div>
              <div style={styles.rowButtonsBottom}>
                <button onClick={() => runView("landskap", "Landskap", buildCurrentConfig())} disabled={!landscapeGraphEnabled} style={!landscapeGraphEnabled ? { ...styles.secondaryButton, ...styles.disabledButton } : styles.secondaryButton}>Landskap</button>
                <button onClick={() => runView("kommun", "Kommun", buildCurrentConfig())} disabled={!municipalityGraphEnabled} style={!municipalityGraphEnabled ? { ...styles.secondaryButton, ...styles.disabledButton } : styles.secondaryButton}>Kommun</button>
                <button onClick={() => runView("karta", "Karta", buildCurrentConfig())} style={styles.secondaryButton}>Karta</button>
              </div>
              <div style={styles.rowButtonsThird}>
                <button onClick={() => runView("fenologi", "Fenologi", buildCurrentConfig())} style={{ ...styles.secondaryButton, gridColumn: 1, gridRow: 1 }}>Fenologi</button>
                <button onClick={() => runView("arter", "Arter", buildCurrentConfig())} disabled={!isMultiSpeciesSelection} style={isMultiSpeciesSelection ? { ...styles.secondaryButton, gridColumn: 1, gridRow: 2 } : { ...styles.secondaryButton, ...styles.disabledButton, gridColumn: 1, gridRow: 2 }}>Arter</button>
                <button onClick={updateLocalDb} disabled={localDbInfo.inProgress || !localDbInfo.exists} style={{ gridColumn: 2, gridRow: 1, ...((localDbInfo.inProgress || !localDbInfo.exists) ? { ...styles.secondaryButton, ...styles.disabledButton } : { ...styles.secondaryButton, background: "#16a34a", borderColor: "#16a34a", color: "white" }) }}>Uppdatera databas</button>
                <button onClick={createLocalDb} disabled={localDbInfo.inProgress} style={{ gridColumn: 2, gridRow: 2, ...(localDbInfo.inProgress ? { ...styles.secondaryButton, ...styles.disabledButton, background: "#fca5a5", borderColor: "#dc2626" } : { ...styles.secondaryButton, background: "#dc2626", borderColor: "#dc2626" }) }}>Skapa lokal databas</button>
              </div>
            </div>
          </div>

          <div style={styles.rightCard}>
            <div style={styles.headerLine}>
              <div>
                <div style={{ fontSize: 22, fontWeight: 700 }}>
                  {viewModeLabel}
                </div>
                <div style={{ fontSize: 14, color: "#475569" }}>Antal visade rader: {resultCount}</div>
              </div>
              <div style={styles.helperText}>
                {voiceEnabled && viewMode === "lista" ? 'Säg "Sortera" + Kolumnrubrik' : voiceEnabled && viewMode === "karta" ? 'Säg "Zooma" + Landskap\n"Zooma" + Kommun\n"Zooma ut"' : ""}
              </div>
              <button onClick={() => (exportIsExcel ? exportToExcel() : exportToPdf())} disabled={!currentReportHasContent} style={exportStyle}>{exportLabel}</button>
            </div>

            <div ref={reportRef} style={{ background: "#ffffff", borderRadius: 14 }}>
            {viewMode === "lista" ? (
              <div style={styles.tableWrap}>
                <div style={{ ...styles.scrollArea, maxHeight: dynamicScrollMaxHeight }}>
                  {results.length === 0 ? <div style={styles.empty}>Ingen lista skapad ännu.</div> : (
                    <table style={styles.table}>
                      <thead><tr>{columns.map((column) => <th key={column.key} style={{ ...styles.th, width: column.width }} onClick={() => handleSort(column.key)}>{column.label}{sortField === column.key ? (sortDirection === "asc" ? " ▲" : " ▼") : ""}</th>)}</tr></thead>
                      <tbody>{results.map((row) => <tr key={row.occurrenceId || `${row.species}-${row.date}-${row.municipality}-${row.locality}`}>{columns.map((column) => <td key={column.key} style={{ ...styles.td, width: column.width }} title={speciesAwareCell(row, column.key) || ""}>{speciesAwareCell(row, column.key) || "—"}</td>)}</tr>)}</tbody>
                    </table>
                  )}
                </div>
              </div>
            ) : viewMode === "fenologi" ? (
              <div style={styles.tableWrap}>
                <div style={{ ...styles.scrollArea, maxHeight: dynamicScrollMaxHeight }}>
                  {phenologyResults.length === 0 ? <div style={styles.empty}>Ingen fenologirapport skapad ännu.</div> : (
                    <table style={styles.table}>
                      <thead><tr>{columns.map((column) => <th key={column.key} style={{ ...styles.th, width: column.width }}>{column.label}</th>)}</tr></thead>
                      <tbody>
                        <tr><td colSpan={columns.length} style={styles.phenologyMarkerCell}>Tidigast</td></tr>
                        {phenologyResults.filter((row) => row.__phenologySection === "tidigast").map((row) => (
                          <tr key={row.occurrenceId || `${row.species}-${row.date}-${row.municipality}-${row.locality}-tidigast`}>
                            {columns.map((column) => <td key={column.key} style={{ ...styles.td, width: column.width }} title={speciesAwareCell(row, column.key) || ""}>{speciesAwareCell(row, column.key) || "—"}</td>)}
                          </tr>
                        ))}
                        <tr><td colSpan={columns.length} style={styles.phenologyMarkerCell}>Senast</td></tr>
                        {phenologyResults.filter((row) => row.__phenologySection === "senast").map((row) => (
                          <tr key={row.occurrenceId || `${row.species}-${row.date}-${row.municipality}-${row.locality}-senast`}>
                            {columns.map((column) => <td key={column.key} style={{ ...styles.td, width: column.width }} title={speciesAwareCell(row, column.key) || ""}>{speciesAwareCell(row, column.key) || "—"}</td>)}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            ) : viewMode === "graf" || viewMode === "tid" ? (
              <div style={styles.chartWrap}>
                {activeChartBlocks.length === 0 ? <div style={styles.empty}>Ingen graf skapad ännu.</div> : (
                  <svg viewBox={`0 0 ${chartWidth} ${chartHeight}`} style={styles.chartSvg}>
                    <line x1={chartLeft} y1={chartTop} x2={chartLeft} y2={chartHeight - chartBottom} stroke="#666" strokeWidth="1" />
                    <line x1={chartLeft} y1={chartHeight - chartBottom} x2={chartWidth - 10} y2={chartHeight - chartBottom} stroke="#666" strokeWidth="1" />
                    {Array.from({ length: Math.floor(yMax / yStep) + 1 }, (_, i) => i * yStep).map((value) => { const y = chartHeight - chartBottom - (value / yMax) * plotHeight; return <g key={value}><line x1={chartLeft} y1={y} x2={chartWidth - 10} y2={y} stroke="#d1d5db" strokeWidth="1" /><text x={chartLeft - 8} y={y + 5} fontSize="14" textAnchor="end" fill="#334155">{value}</text></g>; })}
                    <text x={16} y={12} fontSize="16" fill="#334155">Antal</text>
                    {Array.from({ length: monthCount }, (_, monthGroupIndex) => {
                      const firstBlock = activeChartBlocks[monthGroupIndex * 4];
                      if (!firstBlock) return null;
                      const groupStartX = chartLeft + 8 + monthGroupIndex * groupWidth;
                      const groupEndX = groupStartX + groupWidth - 8;
                      const centerX = groupStartX + 1.5 * (barWidth + barGap) + barWidth / 2;
                      const yearStart = monthGroupIndex === 0 || activeChartBlocks[(monthGroupIndex - 1) * 4]?.year !== firstBlock.year;
                      const yearBg = firstBlock.year % 2 === 0 ? "rgba(0,0,0,0.03)" : "transparent";
                      return <g key={`${firstBlock.year ?? "any"}-${firstBlock.monthIndex}`}><rect x={groupStartX - 4} y={chartTop} width={groupWidth} height={plotHeight} fill={viewMode === "graf" ? yearBg : "transparent"} /><line x1={groupStartX - 4} y1={chartTop} x2={groupStartX - 4} y2={chartHeight - chartBottom} stroke="#cbd5e1" strokeWidth="1" /><text x={centerX} y={chartTop - 8} fontSize="15" textAnchor="middle" fill="#334155">{firstBlock.monthLabel}</text>{viewMode === "graf" && yearStart ? <text x={centerX} y={chartHeight - chartBottom + 30} fontSize="13" textAnchor="middle" fill="#475569">{firstBlock.year}</text> : null}{monthGroupIndex === monthCount - 1 ? <line x1={groupEndX} y1={chartTop} x2={groupEndX} y2={chartHeight - chartBottom} stroke="#cbd5e1" strokeWidth="1" /> : null}</g>;
                    })}
                    {activeChartBlocks.map((block, index) => {
                      const monthGroupIndex = Math.floor(index / 4);
                      const withinMonth = index % 4;
                      const groupStartX = chartLeft + 8 + monthGroupIndex * groupWidth;
                      const x = groupStartX + withinMonth * (barWidth + barGap);
                      const barHeight = (block.count / yMax) * plotHeight;
                      const y = chartHeight - chartBottom - barHeight;
                      return <rect key={block.key} x={x} y={y} width={barWidth} height={barHeight} fill="#355d8a" />;
                    })}
                  </svg>
                )}
              </div>
            ) : viewMode === "landskap" || viewMode === "kommun" || viewMode === "arter" ? (
              <div style={styles.chartWrap}>
                {horizontalSeries.length === 0 ? <div style={styles.empty}>Ingen graf skapad ännu.</div> : (
                  <svg viewBox={`0 0 ${horizontalChartWidth} ${horizontalChartHeight}`} style={styles.chartSvg}>
                    {horizontalSeries.map((item, index) => {
                      const y = horizontalTop + index * 30;
                      const width = (item.count / horizontalMax) * horizontalPlotWidth;
                      return <g key={item.name}><text x={horizontalLeft - 8} y={y + 16} fontSize="13" textAnchor="end" fill="#334155">{item.name}</text><rect x={horizontalLeft} y={y + 2} width={width} height={20} fill="#355d8a" /><text x={horizontalLeft + width + 6} y={y + 16} fontSize="13" fill="#334155">{item.count}</text></g>;
                    })}
                  </svg>
                )}
              </div>
            ) : (
              <div style={styles.mapWrap}>
                {mapPoints.length === 0 ? <div style={styles.empty}>Ingen karta skapad ännu.</div> : (
                  <MapContainer center={mapCenter} zoom={6} scrollWheelZoom={true} preferCanvas={true} style={{ height: `${dynamicMapHeight}px`, width: "100%" }}>
                    <MapViewportController targetBounds={mapTargetBounds} viewportToken={mapViewportToken} mapRefExternal={mapRef} />
                    <TileLayer url="https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png" attribution="&copy; OpenStreetMap" crossOrigin="anonymous" />
                    {mapPoints.map((row) => (
                      <CircleMarker key={row.occurrenceId || `${row.latitude}-${row.longitude}-${row.date}`} center={[row.latitude, row.longitude]} radius={5} pathOptions={{ color: "#b91c1c", fillColor: "#ef4444", fillOpacity: 0.85 }}>
                        <Popup>
                          <div><strong>{row.species}</strong></div>
                          <div>{row.date}</div>
                          <div>{row.locality}</div>
                          <div>{row.municipality}, {row.province}</div>
                          <div>Observatör: {row.recordedBy}</div>
                        </Popup>
                      </CircleMarker>
                    ))}
                  </MapContainer>
                )}
              </div>
            )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}
