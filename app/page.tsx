'use client';

import React, { useEffect, useState, useMemo, Suspense } from 'react';
import { useSearchParams, useRouter } from 'next/navigation';
import { supabase } from '../lib/supabase';
import { motion, AnimatePresence } from 'framer-motion';
import { 
  BookOpen, ChevronLeft, ChevronRight, Info, Search, X, ListOrdered, 
  ArrowRight, Sparkles, FileText, Table as TableIcon, BarChart3, 
  Download, Printer, Copy, Check, Maximize2, Minimize2, Share2, Calendar
} from 'lucide-react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer, 
  PieChart, Pie, Cell, LineChart, Line, AreaChart, Area, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar,
  ComposedChart, LabelList, RadialBarChart, RadialBar, FunnelChart, Funnel, Treemap, Sector, ScatterChart, Scatter
} from 'recharts';

import { ComposableMap, Geographies, Geography, ZoomableGroup } from 'react-simple-maps';
import { scaleLinear } from 'd3-scale';

// LISTA COMPLETA DE FUENTES ESTILO MICROSOFT WORD PARA RENDERIZADO PÚBLICO
const NOMBRES_FUENTES_WORD: Record<string, { label: string; family: string }> = {
  'arial': { label: 'Arial', family: 'Arial, Helvetica, sans-serif' },
  'arial-black': { label: 'Arial Black', family: '"Arial Black", Gadget, sans-serif' },
  'bahnschrift': { label: 'Bahnschrift', family: 'Bahnschrift, "DIN Alternate", sans-serif' },
  'calibri': { label: 'Calibri', family: 'Calibri, Candara, Segoe, "Segoe UI", Optima, Arial, sans-serif' },
  'cambria': { label: 'Cambria', family: 'Cambria, Georgia, serif' },
  'candara': { label: 'Candara', family: 'Candara, Calibri, Segoe, "Segoe UI", Optima, Arial, sans-serif' },
  'century-gothic': { label: 'Century Gothic', family: '"Century Gothic", CenturyGothic, AppleGothic, sans-serif' },
  'comic-sans-ms': { label: 'Comic Sans MS', family: '"Comic Sans MS", cursive, sans-serif' },
  'consolas': { label: 'Consolas', family: 'Consolas, monaco, monospace' },
  'constantia': { label: 'Constantia', family: 'Constantia, "Lucida Bright", Lucidabright, "Lucida Serif", Lucida, "DejaVu Serif", "Bitstream Vera Serif", "Liberation Serif", Georgia, serif' },
  'corbel': { label: 'Corbel', family: 'Corbel, "Lucida Grande", "Lucida Sans Unicode", "Lucida Sans", "DejaVu Sans", "Liberation Sans", Verdana, "Verdana Ref", sans-serif' },
  'courier-new': { label: 'Courier New', family: '"Courier New", Courier, monospace' },
  'franklin-gothic-medium': { label: 'Franklin Gothic', family: '"Franklin Gothic Medium", "Franklin Gothic", "ITC Franklin Gothic", Arial, sans-serif' },
  'garamond': { label: 'Garamond', family: 'Garamond, Baskerville, "Baskerville Old Face", "Hoefler Text", "Times New Roman", serif' },
  'georgia': { label: 'Georgia', family: 'Georgia, serif' },
  'impact': { label: 'Impact', family: 'Impact, Haettenschweiler, "Franklin Gothic Bold", Charcoal, "Helvetica Inserat", "Bitstream Vera Sans Bold", "Arial Black", sans-serif' },
  'lucida-console': { label: 'Lucida Console', family: '"Lucida Console", "Lucida Sans Typewriter", monaco, "Bitstream Vera Sans Mono", monospace' },
  'lucida-sans-unicode': { label: 'Lucida Sans', family: '"Lucida Sans Unicode", "Lucida Grande", sans-serif' },
  'palatino-linotype': { label: 'Palatino Linotype', family: '"Palatino Linotype", "Book Antiqua", Palatino, serif' },
  'segoe-ui': { label: 'Segoe UI', family: '"Segoe UI", Frutiger, "Frutiger Linotype", "Dejavu Sans", "Helvetica Neue", Arial, sans-serif' },
  'tahoma': { label: 'Tahoma', family: 'Tahoma, Geneva, Verdana, sans-serif' },
  'times-new-roman': { label: 'Times New Roman', family: '"Times New Roman", Times, Baskerville, Georgia, serif' },
  'trebuchet-ms': { label: 'Trebuchet MS', family: '"Trebuchet MS", "Lucida Grande", "Lucida Sans Unicode", "Lucida Sans", Tahoma, sans-serif' },
  'verdana': { label: 'Verdana', family: 'Verdana, Geneva, sans-serif' }
};

// CATÁLOGO DE PALETAS DE COLORES PARA LAS GRÁFICAS (COLORES SECUNDARIOS POR DEFECTO)
const PALETAS_GRAFICAS: Record<string, { nombre: string; colores: string[] }> = {
  secundarios: {
    nombre: 'Colores Secundarios (Por Defecto)',
    colores: [
      '#f97316', '#8b5cf6', '#10b981', '#eab308', '#06b6d4',
      '#ec4899', '#14b8a6', '#f43f5e', '#6366f1', '#84cc16',
      '#d97706', '#a855f7', '#059669', '#0284c7', '#e11d48'
    ]
  },
  vivos: {
    nombre: 'Colores Vivos / Vibrantes',
    colores: [
      '#2563eb', '#dc2626', '#16a34a', '#ca8a04', '#9333ea',
      '#0891b2', '#ea580c', '#db2777', '#4f46e5', '#65a30d',
      '#0284c7', '#b91c1c', '#15803d', '#a16207', '#7e22ce'
    ]
  },
  pasteles: {
    nombre: 'Colores Pasteles Suaves',
    colores: [
      '#93c5fd', '#86efac', '#fda4af', '#d8b4fe', '#fdba74',
      '#5eead4', '#f9a8d4', '#fde047', '#a5b4fc', '#cbd5e1',
      '#bef264', '#67e8f9', '#fca5a5', '#c4b5fd', '#fed7aa'
    ]
  },
  acofi_institucional: {
    nombre: 'Institucional ACOFI & Académico',
    colores: [
      '#1e3a8a', '#0284c7', '#0d9488', '#d97706', '#475569',
      '#2563eb', '#0369a1', '#0f766e', '#b45309', '#64748b',
      '#1d4ed8', '#38bdf8', '#2dd4bf', '#f59e0b', '#94a3b8'
    ]
  },
  oceano: {
    nombre: 'Océano y Esmeralda',
    colores: [
      '#0f172a', '#1e40af', '#0369a1', '#0891b2', '#0d9488',
      '#059669', '#10b981', '#34d399', '#38bdf8', '#60a5fa',
      '#818cf8', '#2dd4bf', '#22d3ee', '#93c5fd', '#a7f3d0'
    ]
  },
  atardecer: {
    nombre: 'Atardecer Cálido (Terracota & Oro)',
    colores: [
      '#9a3412', '#c2410c', '#ea580c', '#f97316', '#f59e0b',
      '#eab308', '#dc2626', '#e11d48', '#be185d', '#9d174d',
      '#b45309', '#d97706', '#fb923c', '#fbbf24', '#f43f5e'
    ]
  },
  bosque: {
    nombre: 'Bosque y Naturaleza',
    colores: [
      '#14532d', '#166534', '#15803d', '#16a34a', '#22c55e',
      '#4d7c0f', '#65a30d', '#84cc16', '#0f766e', '#0d9488',
      '#a16207', '#ca8a04', '#10b981', '#34d399', '#86efac'
    ]
  },
  monocromatico: {
    nombre: 'Ejecutivo Sobrio (Grises & Azul Acero)',
    colores: [
      '#0f172a', '#1e293b', '#334155', '#475569', '#64748b',
      '#1e3a8a', '#1d4ed8', '#3b82f6', '#94a3b8', '#cbd5e1',
      '#27272a', '#52525b', '#71717a', '#a1a1aa', '#d4d4d8'
    ]
  }
};

const ESTILOS_FUENTES_PUBLICO = Object.keys(NOMBRES_FUENTES_WORD).map(f => `
  .ql-Font-${f}, .ql-font-${f} { font-family: ${NOMBRES_FUENTES_WORD[f].family}; }
`).join('\n') + `
  .ql-size-small { font-size: 0.75em; }
  .ql-size-large { font-size: 1.5em; }
  .ql-size-huge { font-size: 2.25em; }
  .ql-align-center { text-align: center; }
  .ql-align-center img { margin-left: auto !important; margin-right: auto !important; display: block !important; }
  .ql-align-right { text-align: right; }
  .ql-align-right img { margin-left: auto !important; margin-right: 0 !important; display: block !important; }
  .ql-align-justify { text-align: justify; }
  a[data-inline-btn="true"]:hover {
    transform: scale(1.05) !important;
    filter: brightness(1.06);
  }
  .prose {
    max-width: 100% !important;
    overflow-wrap: break-word !important;
    word-wrap: break-word !important;
    word-break: normal !important;
    white-space: normal !important;
  }
  .prose a:not([data-inline-btn="true"]):not([data-custom-link="true"]) {
    color: #1d4ed8;
    font-weight: 700;
    transition: opacity 0.2s ease;
  }
  .prose a[data-custom-link="true"] {
    transition: opacity 0.2s ease;
  }
  .prose a[data-bold="true"] {
    font-weight: 700 !important;
  }
  .prose a[data-bold="false"] {
    font-weight: 400 !important;
  }
  .prose a[data-italic="true"] {
    font-style: italic !important;
  }
  .prose a[data-italic="false"] {
    font-style: normal !important;
  }
  .prose a[data-underline="true"][data-strike="true"] {
    text-decoration: underline line-through !important;
  }
  .prose a[data-underline="true"]:not([data-strike="true"]) {
    text-decoration: underline !important;
  }
  .prose a[data-underline="false"][data-strike="true"] {
    text-decoration: line-through !important;
  }
  .prose a[data-underline="false"]:not([data-strike="true"]) {
    text-decoration: none !important;
  }
  .prose a:hover {
    opacity: 0.8;
  }
  .prose ul {
    list-style-type: disc !important;
    list-style-position: outside !important;
    padding-left: 1.75rem !important;
    margin-top: 0.5rem !important;
    margin-bottom: 0.75rem !important;
  }
  .prose ol {
    list-style-type: decimal !important;
    list-style-position: outside !important;
    padding-left: 1.75rem !important;
    margin-top: 0.5rem !important;
    margin-bottom: 0.75rem !important;
  }
  .prose li {
    display: list-item !important;
    margin-bottom: 0.25rem !important;
  }
  .prose li[data-list="bullet"] {
    list-style-type: disc !important;
  }
  .prose li[data-list="ordered"] {
    list-style-type: decimal !important;
  }
  .prose .ql-ui {
    display: none !important;
  }
  .prose .ql-indent-1 { padding-left: 3em !important; }
  .prose .ql-indent-2 { padding-left: 6em !important; }
  .prose .ql-indent-3 { padding-left: 9em !important; }
  .prose img {
    margin-top: 0.5rem;
    margin-bottom: 0.5rem;
    max-width: 100% !important;
    height: auto !important;
  }
  .prose img[data-borde="cuadrado"], img[data-borde="cuadrado"] {
    border-radius: 0px !important;
  }
  .prose img[data-borde="redondeado"], img[data-borde="redondeado"] {
    border-radius: 14px !important;
  }
  .prose img[data-borde="circular"], img[data-borde="circular"] {
    border-radius: 9999px !important;
  }
  .prose img[data-alineacion="izquierda"], img[data-alineacion="izquierda"] {
    display: block !important;
    margin-left: 0 !important;
    margin-right: auto !important;
  }
  .prose img[data-alineacion="centro"], img[data-alineacion="centro"] {
    display: block !important;
    margin-left: auto !important;
    margin-right: auto !important;
  }
  .prose img[data-alineacion="derecha"], img[data-alineacion="derecha"] {
    display: block !important;
    margin-left: auto !important;
    margin-right: 0 !important;
  }
  .prose p, .prose div {
    margin-top: 0 !important;
    margin-bottom: 0.35rem !important;
    min-height: 1.5em;
    line-height: 1.6 !important;
    max-width: 100% !important;
    overflow-wrap: break-word !important;
    word-wrap: break-word !important;
    white-space: normal !important;
  }
  .prose h1, .prose h2, .prose h3, .prose h4 {
    margin-top: 0.5rem !important;
    margin-bottom: 0.5rem !important;
    overflow-wrap: break-word !important;
    white-space: normal !important;
  }
  @media print {
    @page {
      size: A4 portrait;
      margin: 10mm 12mm;
    }
    html, body {
      height: auto !important;
      overflow: visible !important;
      -webkit-print-color-adjust: exact !important;
      print-color-adjust: exact !important;
      background: #fdfbf7 !important;
    }
    .no-print {
      display: none !important;
    }
    .print-book-wrapper {
      display: block !important;
      position: static !important;
      width: 100% !important;
      height: auto !important;
      overflow: visible !important;
    }
    .print-page-sheet {
      page-break-before: always !important;
      break-before: page !important;
      width: 100% !important;
      padding: 6mm 4mm !important;
      box-sizing: border-box !important;
      position: relative !important;
      overflow: visible !important;
    }
    .print-page-sheet:first-child {
      page-break-before: auto !important;
      break-before: auto !important;
    }
    .print-header-keep {
      page-break-after: avoid !important;
      break-after: avoid !important;
      page-break-inside: avoid !important;
      break-inside: avoid !important;
    }
    .print-avoid-break {
      page-break-inside: avoid !important;
      break-inside: avoid !important;
    }
  }
`;

type Edicion = { id: number; titulo: string; anio: string; color_principal: string; estado: string; contenido_presentacion: string; };

type BloqueBase = { id: string; colorFondo?: string; escala?: number; };

type BloqueTexto = BloqueBase & { 
  tipo: 'texto'; 
  contenido: string; 
  columnas?: 1 | 2 | 3; 
  contenidoCol2?: string; 
  contenidoCol3?: string; 
};

type ItemFirma = {
  id: string;
  url: string;
  nombre: string;
  cargo: string;
  infoAdicional?: string;
  tamanoFirma?: number;
  espacioSuperior?: number;
};

type BloqueFirma = BloqueBase & {
  tipo: 'firma';
  alineacion: 'izquierda' | 'centro' | 'derecha' | 'espaciado';
  espacioGeneralSuperior?: number;
  espacioGeneralInferior?: number;
  firmas: ItemFirma[];
};

type BloqueTabla = BloqueBase & { 
  tipo: 'tabla'; 
  nombre: string; 
  data: string[][]; 
  oculto?: boolean; 
  coloresFilas?: Record<number, string>;
  coloresColumnas?: Record<number, string>;
  paddingFila?: number;
  tamanoTextoTabla?: number;
  anchosColumnas?: Record<number, number>;
  altosFilas?: Record<number, number>;
  pieBloque?: string;
};

type TipoGrafica = 
  | 'barras' 
  | 'barras_horizontales' 
  | 'barras_apiladas' 
  | 'barras_4d' 
  | 'barras_3d_apiladas'
  | 'torta' 
  | 'torta_rellena'
  | 'torta_3d_rellena'
  | 'rosa_polar'
  | 'dona' 
  | 'torta_4d' 
  | 'semicirculo'
  | 'lineas' 
  | 'lineas_curvas'
  | 'area' 
  | 'area_apilada'
  | 'radar' 
  | 'radar_relleno'
  | 'barras_radiales'
  | 'embudo'
  | 'treemap'
  | 'dispersion'
  | 'combinada' 
  | 'mapa_colombia';

type BloqueGrafica = BloqueBase & { 
  tipo: 'grafica'; 
  nombre: string; 
  tablaRef: string; 
  tipoGrafica: TipoGrafica; 
  ejeX: number; 
  ejesY: number[]; 
  coloresCustom?: string[]; 
  paletaNombre?: string;
  mostrarValores?: boolean; 
  ocultarEjeX?: boolean;
  ocultarLeyenda?: boolean;
  alturaGrafica?: number;
  grosorBarra?: number;
  espacioEntreBarras?: number;
  ejeYMin?: number | string;
  ejeYMax?: number | string;
  pieBloque?: string;
};

type BloqueImagen = BloqueBase & { tipo: 'imagen'; url: string; alt: string; pantallaCompleta: boolean; };

type EstiloBoton = 'solido' | 'contorno' | 'pildora' | '3d' | '4d';
type TamanoBoton = 'pequeno' | 'mediano' | 'grande' | 'extragrande';
type BloqueBoton = BloqueBase & {
  tipo: 'boton';
  texto: string;
  url: string;
  estilo: EstiloBoton;
  tamanoBoton?: TamanoBoton;
  fuenteBoton?: string;
  colorBoton: string;
  colorTexto: string;
  colorIluminacion?: string;
  alineacion: 'izquierda' | 'centro' | 'derecha';
  abrirNuevaPestana?: boolean;
};

type Bloque = BloqueTexto | BloqueFirma | BloqueTabla | BloqueGrafica | BloqueImagen | BloqueBoton;

type Pagina = { 
  id: string; 
  titulo: string; 
  subtitulo?: string; 
  mostrarNumero?: boolean; 
  posicion?: 'antes_portada' | 'despues_portada'; 
  cabeceraHtml?: string;
  piePaginaHtml?: string;
  bloques: Bloque[]; 
  colorFondo?: string; 
  imagenFondo?: string;
  opacidadFondo?: number;
};

type ConfiguracionEdicion = {
  subtitulo?: string;
  portadaContenido?: string;
  colorFondoPortada?: string;
  colorFondoPaginasGlobal?: string;
  imagenFondoPortada?: string;
  portadaBgBrightness?: number;
  portadaBgContrast?: number;
  portadaBgOverlayOpacity?: number;
  portadaBgBlur?: number;
  logoUrl?: string;
  logoSize?: number;
  logoBrightness?: number;
  marcaAguaUrl?: string;
  marcaAguaSize?: number;
  marcaAguaOpacity?: number;
  modoMarcaAgua?: 'todas' | 'seleccionadas';
  paginasMarcaAgua?: string[];
  cabeceraGlobalHtml?: string;
  piePaginaGlobalHtml?: string;
  mostrarCabeceraGlobal?: boolean;
  mostrarPiePaginaGlobal?: boolean;
};

const quitarAcentos = (str: string) =>
  (str || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();

const limpiarHtmlYEntidades = (html?: string) => {
  if (!html) return '';
  return html
    .replace(/<\/(p|div|h[1-6]|li|tr|br)>/gi, ' ')
    .replace(/<br\s*\/?>/gi, ' ')
    .replace(/<[^>]*>/g, ' ')
    .replace(/&nbsp;|&#160;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&ldquo;|&rdquo;|&laquo;|&raquo;/gi, '"')
    .replace(/&ndash;|&mdash;/gi, '-')
    .replace(/&aacute;/gi, 'á')
    .replace(/&eacute;/gi, 'é')
    .replace(/&iacute;/gi, 'í')
    .replace(/&oacute;/gi, 'ó')
    .replace(/&uacute;/gi, 'ú')
    .replace(/&ntilde;/gi, 'ñ')
    .replace(/&Aacute;/gi, 'Á')
    .replace(/&Eacute;/gi, 'É')
    .replace(/&Iacute;/gi, 'Í')
    .replace(/&Oacute;/gi, 'Ó')
    .replace(/&Uacute;/gi, 'Ú')
    .replace(/&Ntilde;/gi, 'Ñ')
    .replace(/\[\[BOTON\|([^\]]*)\]\]/gi, (_, inner) => String(inner).split('|')[0] || '')
    .replace(/\s+/g, ' ')
    .trim();
};

function LibroPublicoContent() {
  const searchParams = useSearchParams();
  const router = useRouter();
  const parametroEdicion = searchParams?.get('edicion');
  const parametroPag = searchParams?.get('pag');

  const [edicionesDisponibles, setEdicionesDisponibles] = useState<{ id: number; anio: string; titulo: string }[]>([]);
  const [edicionActiva, setEdicionActiva] = useState<Edicion | null>(null);
  const [paginas, setPaginas] = useState<Pagina[]>([]);
  const [configuracion, setConfiguracion] = useState<ConfiguracionEdicion>({});
  const [cargando, setCargando] = useState(true);
  
  const [paginaActual, setPaginaActual] = useState(0);
  const [direccionTransicion, setDireccionTransicion] = useState<1 | -1>(1);
  const [memoriaInicializada, setMemoriaInicializada] = useState(false);
  
  const [tooltipMapa, setTooltipMapa] = useState<{ nombre: string; valor: string; x: number; y: number } | null>(null);

  const [colombiaGeoJson, setColombiaGeoJson] = useState<any>(null);

  const [buscadorAbierto, setBuscadorAbierto] = useState(false);
  const [terminoBusqueda, setTerminoBusqueda] = useState('');
  const [filtroCategoriaBusqueda, setFiltroCategoriaBusqueda] = useState<'todo' | 'capitulos' | 'tablas'>('todo');
  const [indiceResultadoSeleccionado, setIndiceResultadoSeleccionado] = useState(0);
  const [inputIrAPagina, setInputIrAPagina] = useState('');

  const [generandoPdfPublico, setGenerandoPdfPublico] = useState(false);
  const [filtrosTablas, setFiltrosTablas] = useState<Record<string, string>>({});
  const [tablaCopiadaId, setTablaCopiadaId] = useState<string | null>(null);
  const [enlaceCompartido, setEnlaceCompartido] = useState(false);
  const [esPantallaCompleta, setEsPantallaCompleta] = useState(false);

  const [touchStart, setTouchStart] = useState<{ x: number, y: number } | null>(null);
  const [touchEnd, setTouchEnd] = useState<{ x: number, y: number } | null>(null);
  const minSwipeDistance = 50; 

  const COLORES_GRAFICAS_DEFAULT = PALETAS_GRAFICAS.secundarios.colores;

  useEffect(() => {
    fetch('/colombia.json')
      .then(res => res.json())
      .then(data => setColombiaGeoJson(data))
      .catch(() => {});
  }, []);

  useEffect(() => {
    const cargarDatos = async () => {
      try {
        const { data: listaEds } = await supabase
          .from('ediciones')
          .select('id, anio, titulo, estado')
          .eq('estado', 'publicado')
          .order('anio', { ascending: false });

        if (listaEds) {
          setEdicionesDisponibles(listaEds);
        }

        let query = supabase.from('ediciones').select('*');
        query = parametroEdicion 
          ? query.eq('anio', parametroEdicion) 
          : query.eq('estado', 'publicado').order('anio', { ascending: false }).limit(1);
          
        const { data, error } = await query.single();
        if (error) throw error;
        
        if (data) {
          setEdicionActiva(data as Edicion);
          try {
            const parsed = JSON.parse(data.contenido_presentacion);
            if (Array.isArray(parsed)) {
              setPaginas(parsed);
            } else if (parsed && parsed.paginas) {
              setPaginas(parsed.paginas);
              setConfiguracion(parsed.configuracion || {});
            }
          } catch (e) {
            setPaginas([]);
          }
        }
      } catch (err) {
        console.error("No se pudo cargar la edición:", err);
      } finally {
        setCargando(false);
      }
    };
    cargarDatos();
  }, [parametroEdicion]);

  type SeccionLibro = 
    | { tipo: 'portada' }
    | { tipo: 'pagina'; pagina: Pagina; numeroReal: number };

  const secciones: SeccionLibro[] = useMemo(() => {
    const paginasAntes = paginas.filter(p => p.posicion === 'antes_portada');
    const paginasDespues = paginas.filter(p => p.posicion !== 'antes_portada');
    return [
      ...paginasAntes.map((p, idx) => ({ tipo: 'pagina' as const, pagina: p, numeroReal: idx + 1 })),
      { tipo: 'portada' as const },
      ...paginasDespues.map((p, idx) => ({ tipo: 'pagina' as const, pagina: p, numeroReal: paginasAntes.length + idx + 1 }))
    ];
  }, [paginas]);

  const totalSecciones = secciones.length;

  const navegarASeccion = (nuevoIdx: number) => {
    const idxSeguro = Math.max(0, Math.min(totalSecciones - 1, nuevoIdx));
    if (idxSeguro === paginaActual) return;
    setDireccionTransicion(idxSeguro > paginaActual ? 1 : -1);
    setTooltipMapa(null);
    setPaginaActual(idxSeguro);
  };

  const obtenerEstilosTamanoBoton = (tamano?: TamanoBoton) => {
    switch (tamano) {
      case 'pequeno':
        return { padding: '8px 18px', fontSize: '11px' };
      case 'grande':
        return { padding: '16px 38px', fontSize: '15px' };
      case 'extragrande':
        return { padding: '20px 48px', fontSize: '18px' };
      case 'mediano':
      default:
        return { padding: '13px 28px', fontSize: '13px' };
    }
  };

  const descargarTablaCsv = (tabla: BloqueTabla) => {
    const filasCsv = (tabla.data || []).map(fila =>
      fila.map(celda => `"${String(celda ?? '').replace(/"/g, '""')}"`).join(';')
    ).join('\r\n');
    const blob = new Blob(['\uFEFF' + filasCsv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${(tabla.nombre || 'Tabla_ACOFI').replace(/[^a-zA-Z0-9_-]/g, '_')}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const copiarTablaPortapapeles = (tabla: BloqueTabla) => {
    const textoTabulado = (tabla.data || []).map(fila => fila.join('\t')).join('\n');
    navigator.clipboard.writeText(textoTabulado);
    setTablaCopiadaId(tabla.id);
    setTimeout(() => setTablaCopiadaId(null), 2200);
  };

  const compartirPaginaActual = () => {
    if (typeof window === 'undefined') return;
    const url = new URL(window.location.href);
    if (edicionActiva?.anio) url.searchParams.set('edicion', edicionActiva.anio);
    const sec = secciones[paginaActual];
    if (sec && sec.tipo === 'pagina') {
      url.searchParams.set('pag', String(sec.numeroReal));
    } else {
      url.searchParams.delete('pag');
    }
    navigator.clipboard.writeText(url.toString());
    setEnlaceCompartido(true);
    setTimeout(() => setEnlaceCompartido(false), 2500);
  };

  const alternarPantallaCompleta = () => {
    if (typeof document === 'undefined') return;
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen?.().then(() => setEsPantallaCompleta(true)).catch(() => {});
    } else {
      document.exitFullscreen?.().then(() => setEsPantallaCompleta(false)).catch(() => {});
    }
  };

  const descargarLibroCompletoPdf = () => {
    setGenerandoPdfPublico(true);
    setTimeout(() => {
      window.print();
      setTimeout(() => {
        setGenerandoPdfPublico(false);
      }, 900);
    }, 1800);
  };

  const mapaTablasPorId = useMemo(() => {
    const mapa: Record<string, BloqueTabla> = {};
    paginas.forEach(p => {
      p.bloques.forEach(b => {
        if (b.tipo === 'tabla') {
          mapa[b.id] = b as BloqueTabla;
        }
      });
    });
    return mapa;
  }, [paginas]);

  const indiceBusqueda = useMemo(() => {
    return secciones.map((sec, idx) => {
      if (sec.tipo === 'portada') {
        const textoPortada = limpiarHtmlYEntidades(`${edicionActiva?.titulo || ''} ${configuracion.subtitulo || ''} ${configuracion.portadaContenido || ''}`);
        const tituloPortada = edicionActiva?.titulo || 'Portada Principal';
        return {
          idxSeccion: idx,
          esPortada: true,
          numeroReal: 0,
          titulo: tituloPortada,
          tituloNorm: quitarAcentos(tituloPortada),
          subtitulo: 'Portada oficial de la publicación',
          subtituloNorm: 'portada oficial de la publicacion',
          cuerpoLimpio: textoPortada,
          cuerpoNorm: quitarAcentos(textoPortada),
          tieneTabla: false,
          tieneGrafica: false,
          tieneCapitulo: true,
          nombresTablas: ''
        };
      }

      const p = sec.pagina;
      let tieneTabla = false;
      let tieneGrafica = false;
      const nombresTablasArr: string[] = [];

      const fragmentosBloques = p.bloques.map(b => {
        if (b.tipo === 'texto') {
          return `${limpiarHtmlYEntidades(b.contenido)} ${limpiarHtmlYEntidades(b.contenidoCol2)} ${limpiarHtmlYEntidades(b.contenidoCol3)}`;
        }
        if (b.tipo === 'tabla') {
          tieneTabla = true;
          if (b.nombre) nombresTablasArr.push(b.nombre);
          const celdas = (b.data || []).map(f => f.join(' • ')).join(' | ');
          return `${b.nombre} ${celdas} ${limpiarHtmlYEntidades(b.pieBloque)}`;
        }
        if (b.tipo === 'grafica') {
          tieneGrafica = true;
          return `${b.nombre} ${limpiarHtmlYEntidades(b.pieBloque)}`;
        }
        if (b.tipo === 'boton') {
          return b.texto;
        }
        if (b.tipo === 'firma') {
          return (b.firmas || []).map(f => `${f.nombre} ${f.cargo} ${f.infoAdicional || ''}`).join(' ');
        }
        return '';
      }).join(' ');

      const cuerpoLimpio = fragmentosBloques.replace(/\s+/g, ' ').trim();
      const tituloOriginal = p.titulo.trim();
      const subtituloOriginal = (p.subtitulo || '').trim();

      const primerBloqueConNombre = p.bloques.find(b => (b.tipo === 'tabla' || b.tipo === 'grafica') && b.nombre) as (BloqueTabla | BloqueGrafica | undefined);
      const tituloMostrar = tituloOriginal || subtituloOriginal || primerBloqueConNombre?.nombre || `Página ${sec.numeroReal}`;

      return {
        idxSeccion: idx,
        esPortada: false,
        numeroReal: sec.numeroReal,
        titulo: tituloMostrar,
        tituloNorm: quitarAcentos(tituloMostrar),
        subtitulo: subtituloOriginal,
        subtituloNorm: quitarAcentos(subtituloOriginal),
        cuerpoLimpio,
        cuerpoNorm: quitarAcentos(cuerpoLimpio),
        tieneTabla,
        tieneGrafica,
        tieneCapitulo: tituloOriginal !== '',
        nombresTablas: nombresTablasArr.join(' • ')
      };
    });
  }, [secciones, edicionActiva, configuracion]);

  const resultadosBusqueda = useMemo(() => {
    const baseFiltrada = indiceBusqueda.filter(item => {
      if (filtroCategoriaBusqueda === 'capitulos') return item.tieneCapitulo || item.esPortada;
      if (filtroCategoriaBusqueda === 'tablas') return item.tieneTabla || item.tieneGrafica;
      return true;
    });

    const rawQuery = terminoBusqueda.trim();
    if (!rawQuery) {
      return baseFiltrada.slice(0, 30).map(item => ({
        ...item,
        extracto: item.subtitulo || item.nombresTablas || item.cuerpoLimpio.slice(0, 140),
        esCoincidenciaExactaPagina: false,
        terminoResaltar: ''
      }));
    }

    const qNorm = quitarAcentos(rawQuery);
    const matchComandoPagina = qNorm.match(/^(?:pagina|pag\.?|p\.?|hoja|ir a)\s*(\d+)$/) || qNorm.match(/^(\d+)$/);
    const numeroPaginaBuscado = matchComandoPagina ? Number(matchComandoPagina[1]) : null;

    const palabrasCrudas = qNorm.split(/\s+/).filter(Boolean);
    const palabrasSignificativas = palabrasCrudas.filter(
      w => !['pagina', 'pag', 'pag.', 'hoja', 'de', 'del', 'la', 'el', 'en', 'los', 'las'].includes(w)
    );
    const terminosActivos = palabrasSignificativas.length > 0 ? palabrasSignificativas : palabrasCrudas;

    const resultadosConScore: any[] = [];

    for (const item of baseFiltrada) {
      let score = 0;
      let esCoincidenciaExactaPagina = false;

      if (numeroPaginaBuscado !== null) {
        if (!item.esPortada && item.numeroReal === numeroPaginaBuscado) {
          score += 5000;
          esCoincidenciaExactaPagina = true;
        } else if (matchComandoPagina && qNorm.match(/^(?:pagina|pag\.?|p\.?|hoja)\s*\d+$/)) {
          continue;
        }
      }

      if (!esCoincidenciaExactaPagina) {
        const todoNorm = `${item.tituloNorm} ${item.subtituloNorm} ${item.cuerpoNorm}`;
        const cumpleTodos = terminosActivos.every(t => todoNorm.includes(t));
        if (!cumpleTodos) continue;

        if (item.tituloNorm === qNorm) score += 1000;
        else if (item.tituloNorm.includes(qNorm)) score += 600;

        if (item.subtituloNorm.includes(qNorm)) score += 350;
        if (item.cuerpoNorm.includes(qNorm)) score += 150;

        terminosActivos.forEach(t => {
          if (item.tituloNorm.includes(t)) score += 120;
          if (item.subtituloNorm.includes(t)) score += 60;
          if (item.cuerpoNorm.includes(t)) score += 20;
        });
      }

      let extracto = item.subtitulo || item.nombresTablas || '';
      const terminoParaExtracto = terminosActivos[0] || qNorm;
      const posEnCuerpo = item.cuerpoNorm.indexOf(terminoParaExtracto);

      if (posEnCuerpo !== -1) {
        const inicio = Math.max(0, posEnCuerpo - 45);
        const fin = Math.min(item.cuerpoLimpio.length, posEnCuerpo + 115);
        extracto = (inicio > 0 ? '... ' : '') + item.cuerpoLimpio.slice(inicio, fin).trim() + (fin < item.cuerpoLimpio.length ? ' ...' : '');
      } else if (!extracto && item.cuerpoLimpio) {
        extracto = item.cuerpoLimpio.slice(0, 140) + (item.cuerpoLimpio.length > 140 ? ' ...' : '');
      }

      resultadosConScore.push({
        ...item,
        score,
        extracto,
        esCoincidenciaExactaPagina,
        terminoResaltar: terminosActivos[0] || rawQuery
      });
    }

    return resultadosConScore.sort((a, b) => b.score - a.score || a.idxSeccion - b.idxSeccion);
  }, [terminoBusqueda, indiceBusqueda, filtroCategoriaBusqueda]);

  useEffect(() => {
    setIndiceResultadoSeleccionado(0);
  }, [terminoBusqueda, filtroCategoriaBusqueda]);

  const resaltarCoincidencia = (texto: string, termino: string) => {
    if (!texto || !termino || termino.length < 2) return texto;
    const textoSinAcentos = quitarAcentos(texto);
    const terminoSinAcentos = quitarAcentos(termino);
    const idx = textoSinAcentos.indexOf(terminoSinAcentos);
    if (idx === -1) return texto;

    const antes = texto.slice(0, idx);
    const coincidencia = texto.slice(idx, idx + termino.length);
    const despues = texto.slice(idx + termino.length);

    return (
      <>
        {antes}
        <mark className="bg-amber-200/90 text-stone-950 font-bold px-1 rounded">{coincidencia}</mark>
        {despues}
      </>
    );
  };

  const irAPaginaInterna = (identificadorDestino: string) => {
    if (!identificadorDestino) return false;
    const limpio = identificadorDestino.replace(/^#pagina:/i, '').replace(/^#/, '').trim();
    if (!limpio) return false;

    if (limpio.toLowerCase() === 'portada') {
      const idxPortada = secciones.findIndex(s => s.tipo === 'portada');
      if (idxPortada !== -1) {
        navegarASeccion(idxPortada);
        return true;
      }
    }

    const idxPorId = secciones.findIndex(s => s.tipo === 'pagina' && s.pagina.id === limpio);
    if (idxPorId !== -1) {
      navegarASeccion(idxPorId);
      return true;
    }

    const numHoja = Number(limpio);
    if (!isNaN(numHoja) && numHoja >= 1) {
      const idxPorNumero = secciones.findIndex(s => s.tipo === 'pagina' && s.numeroReal === numHoja);
      if (idxPorNumero !== -1) {
        navegarASeccion(idxPorNumero);
        return true;
      }
    }

    const idxPorTitulo = secciones.findIndex(
      s => s.tipo === 'pagina' && quitarAcentos(s.pagina.titulo.trim()) === quitarAcentos(limpio)
    );
    if (idxPorTitulo !== -1) {
      navegarASeccion(idxPorTitulo);
      return true;
    }

    return false;
  };

  const manejarClicContenidoInterno = (e: React.MouseEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement;
    const anchor = target.closest('a');
    if (!anchor) return;

    const attrInternal = anchor.getAttribute('data-internal-page');
    const rawHref = anchor.getAttribute('href') || '';

    if (attrInternal) {
      e.preventDefault();
      irAPaginaInterna(attrInternal);
      return;
    }

    if (rawHref.startsWith('#pagina:') || (rawHref.startsWith('#') && rawHref.length > 1)) {
      e.preventDefault();
      irAPaginaInterna(rawHref);
    }
  };

  useEffect(() => {
    if (edicionActiva && !memoriaInicializada) {
      if (parametroPag) {
        const numPagUrl = Number(parametroPag);
        const idxUrl = secciones.findIndex(s => s.tipo === 'pagina' && s.numeroReal === numPagUrl);
        if (idxUrl !== -1) {
          setPaginaActual(idxUrl);
          setMemoriaInicializada(true);
          return;
        }
      }
      const paginaGuardada = sessionStorage.getItem(`publico_pagina_${edicionActiva.id}`);
      if (paginaGuardada) {
        const indexGuardado = Number(paginaGuardada);
        if (indexGuardado >= 0 && indexGuardado < totalSecciones) {
          setPaginaActual(indexGuardado);
        }
      }
      setMemoriaInicializada(true);
    }
  }, [edicionActiva, totalSecciones, memoriaInicializada, parametroPag, secciones]);

  useEffect(() => {
    if (edicionActiva && memoriaInicializada) {
      sessionStorage.setItem(`publico_pagina_${edicionActiva.id}`, paginaActual.toString());
    }
  }, [paginaActual, edicionActiva, memoriaInicializada]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        setBuscadorAbierto(prev => !prev);
        return;
      }
      if (e.key === 'Escape' && buscadorAbierto) {
        setBuscadorAbierto(false);
        return;
      }
      if (buscadorAbierto) return;
      const tagActivo = (document.activeElement as HTMLElement)?.tagName;
      if (tagActivo === 'INPUT' || tagActivo === 'SELECT' || tagActivo === 'TEXTAREA') return;

      if (e.key === 'ArrowRight' && paginaActual < totalSecciones - 1) navegarASeccion(paginaActual + 1);
      if (e.key === 'ArrowLeft' && paginaActual > 0) navegarASeccion(paginaActual - 1);
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [paginaActual, totalSecciones, buscadorAbierto]);

  const onTouchStart = (e: React.TouchEvent) => {
    setTouchEnd(null);
    setTouchStart({ x: e.targetTouches[0].clientX, y: e.targetTouches[0].clientY });
  };

  const onTouchMove = (e: React.TouchEvent) => {
    setTouchEnd({ x: e.targetTouches[0].clientX, y: e.targetTouches[0].clientY });
  };

  const onTouchEnd = () => {
    if (!touchStart || !touchEnd) return;
    const distanceX = touchStart.x - touchEnd.x;
    const distanceY = touchStart.y - touchEnd.y;
    const isHorizontalSwipe = Math.abs(distanceX) > Math.abs(distanceY);

    if (isHorizontalSwipe && Math.abs(distanceX) > minSwipeDistance) {
      if (distanceX > 0 && paginaActual < totalSecciones - 1) navegarASeccion(paginaActual + 1);
      else if (distanceX < 0 && paginaActual > 0) navegarASeccion(paginaActual - 1);
    }
  };

  if (cargando || (edicionActiva && !memoriaInicializada)) {
    return (
      <div className="flex h-screen items-center justify-center bg-[#fdfbf7] font-serif">
        <div className="w-16 h-16 border-4 border-stone-200 border-t-blue-700 rounded-full animate-spin"></div>
      </div>
    );
  }

  if (!edicionActiva || paginas.length === 0) {
    return (
      <div className="flex flex-col h-screen items-center justify-center bg-[#fdfbf7] text-center p-8 font-serif text-stone-900">
        <BookOpen size={48} className="text-stone-300 mb-6" />
        <h2 className="text-4xl tracking-tight mb-4">Publicación no disponible</h2>
        <p className="text-stone-500 max-w-md">No hay contenido estructurado para esta edición o el enlace es incorrecto.</p>
      </div>
    );
  }

  const parseValor = (val: any) => {
    if (val === undefined || val === null) return 0;
    if (typeof val === 'number') return isFinite(val) ? val : 0;
    let s = String(val).trim();
    if (!s) return 0;
    s = s.replace(/[^0-9.,-]/g, '');
    if (!s || s === '-' || s === '.' || s === ',') return 0;

    if (s.includes('.') && s.includes(',')) {
      if (s.lastIndexOf(',') > s.lastIndexOf('.')) {
        s = s.replace(/\./g, '').replace(/,/g, '.');
      } else {
        s = s.replace(/,/g, '');
      }
    } else if (s.includes(',')) {
      const partes = s.split(',');
      if (partes.length > 2) {
        s = s.replace(/,/g, '');
      } else {
        s = s.replace(',', '.');
      }
    } else if (s.includes('.')) {
      const partes = s.split('.');
      if (partes.length > 2) {
        s = s.replace(/\./g, '');
      } else if (partes[1] && partes[1].length === 3 && Number(partes[0]) >= 1 && !String(val).includes('%')) {
        s = s.replace(/\./g, '');
      }
    }
    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  };

  const obtenerPaletaActivaBloque = (bloqueGrafica: BloqueGrafica) => {
    if (bloqueGrafica.paletaNombre && PALETAS_GRAFICAS[bloqueGrafica.paletaNombre]) {
      return PALETAS_GRAFICAS[bloqueGrafica.paletaNombre].colores;
    }
    return COLORES_GRAFICAS_DEFAULT;
  };

  // FORMATEADOR 100% BLINDADO PARA RECHARTS (ESCAPA PUNTOS "." EN CABECERAS DE EXCEL Y RESPETA EL ORDEN DE SELECCIÓN)
  const formatearDatosParaGrafica = (bloqueGrafica: BloqueGrafica) => {
    const tablaFuente = mapaTablasPorId[bloqueGrafica.tablaRef];
    if (!tablaFuente || !Array.isArray(tablaFuente.data) || tablaFuente.data.length < 2) {
      return { datos: [], cabeceras: [], ejesYUsar: [], sufijoPorcentaje: false, columnasPorcentaje: {} as Record<string, boolean> };
    }

    const cabecerasOriginales = Array.isArray(tablaFuente.data[0]) ? tablaFuente.data[0] : [];
    const totalColumnas = cabecerasOriginales.length;
    if (totalColumnas === 0) {
      return { datos: [], cabeceras: [], ejesYUsar: [], sufijoPorcentaje: false, columnasPorcentaje: {} as Record<string, boolean> };
    }

    const conteoNombres: Record<string, number> = {};
    const cabeceras = cabecerasOriginales.map((c, i) => {
      let limpio = String(c ?? '')
        .replace(/[\r\n\t]+/g, ' ')
        .replace(/\./g, '․')
        .replace(/[\[\]]/g, '')
        .trim();
      if (!limpio) limpio = `Columna ${i + 1}`;
      if (conteoNombres[limpio] !== undefined) {
        conteoNombres[limpio] += 1;
        limpio = `${limpio} (${conteoNombres[limpio]})`;
      } else {
        conteoNombres[limpio] = 1;
      }
      return limpio;
    });

    const filas = tablaFuente.data
      .slice(1)
      .filter(f => Array.isArray(f) && f.some(celda => String(celda ?? '').trim() !== ''));

    if (filas.length === 0) {
      return { datos: [], cabeceras, ejesYUsar: [], sufijoPorcentaje: false, columnasPorcentaje: {} as Record<string, boolean> };
    }

    const rawEjeX = Number(bloqueGrafica.ejeX);
    const ejeXSeguro = !isNaN(rawEjeX) && rawEjeX >= 0 && rawEjeX < totalColumnas ? rawEjeX : 0;

    let ejesYValidos: number[] = [];
    if (Array.isArray(bloqueGrafica.ejesY)) {
      ejesYValidos = Array.from(
        new Set(
          bloqueGrafica.ejesY
            .map(Number)
            .filter(idx => !isNaN(idx) && idx >= 0 && idx < totalColumnas && idx !== ejeXSeguro)
        )
      );
    } else {
      ejesYValidos = cabeceras.map((_, i) => i).filter(i => i !== ejeXSeguro);
    }

    const indiceMapa = ejesYValidos[0] ?? 0;

    let detectaPorcentaje = false;
    const columnasPorcentaje: Record<string, boolean> = {};

    ejesYValidos.forEach(idx => {
      const nombreCab = cabeceras[idx];
      if (!nombreCab) return;
      const cabTienePorcentaje = nombreCab.includes('%') || nombreCab.toLowerCase().includes('porcentaje');
      const filaTienePorcentaje = filas.some(f => String(f?.[idx] ?? '').includes('%'));
      if (cabTienePorcentaje || filaTienePorcentaje) {
        columnasPorcentaje[nombreCab] = true;
        detectaPorcentaje = true;
      }
    });

    const paletaBase = obtenerPaletaActivaBloque(bloqueGrafica);

    const datosFormateados = filas.map((fila, fIdx) => {
      const celdaPrincipal = String(fila?.[indiceMapa] ?? '');
      if (celdaPrincipal.includes('%')) detectaPorcentaje = true;

      const obj: any = { 
        name: String(fila?.[ejeXSeguro] ?? `Ítem ${fIdx + 1}`).trim() || `Ítem ${fIdx + 1}`,
        valorMapa: parseValor(fila?.[indiceMapa]),
        fill: bloqueGrafica.coloresCustom?.[fIdx] || paletaBase[fIdx % paletaBase.length]
      };
      
      ejesYValidos.forEach(idx => {
        const nombreCabecera = cabeceras[idx];
        if (nombreCabecera) {
          obj[nombreCabecera] = parseValor(fila?.[idx]);
        }
      });
      return obj;
    });

    return { datos: datosFormateados, cabeceras, ejesYUsar: ejesYValidos, sufijoPorcentaje: detectaPorcentaje, columnasPorcentaje };
  };

  const normalizarTexto = (texto: any) => {
    if (!texto) return '';
    let t = String(texto).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]/g, '');
    const alias: Record<string, string> = {
      "bogota": "bogotadc", "valle": "valledelcauca", "sanandres": "archipielagodesanandresprovidenciaysantacatalina",
      "sanandresyprovidencia": "archipielagodesanandresprovidenciaysantacatalina", "guajira": "laguajira"
    };
    return alias[t] || t;
  };
  // LEYENDA PERSONALIZADA QUE RESPETA ESTRICTAMENTE EL ORDEN EN QUE SE SELECCIONARON LOS EJES Y
  const renderizarLeyendaEnOrden = (props: any, nombresEjesY: string[], paletaActiva: string[], bloque: BloqueGrafica) => {
    const { payload } = props;
    if (!payload || payload.length === 0) return null;

    return (
      <div className="flex flex-wrap items-center justify-center gap-4 pt-5 font-sans text-xs font-bold text-stone-700">
        {nombresEjesY.map((nombreSerie, idx) => {
          const colorSerie = bloque.coloresCustom?.[idx] || paletaActiva[idx % paletaActiva.length];
          const nombreLimpio = nombreSerie.replace(/․/g, '.');
          return (
            <div key={nombreSerie} className="flex items-center gap-2">
              <span className="w-3.5 h-3.5 rounded-sm shrink-0" style={{ backgroundColor: colorSerie }} />
              <span>{nombreLimpio}</span>
            </div>
          );
        })}
      </div>
    );
  };

  const renderizarGrafica = (bloque: BloqueGrafica) => {
    const { datos, cabeceras, ejesYUsar, sufijoPorcentaje, columnasPorcentaje } = formatearDatosParaGrafica(bloque);
    if (datos.length === 0 || ejesYUsar.length === 0) return <div className="p-8 text-center text-stone-400 italic">Faltan datos en la tabla conectada.</div>;

    const nombresEjesY = ejesYUsar.map(idx => cabeceras[idx]).filter(Boolean);
    if (nombresEjesY.length === 0) return <div className="p-8 text-center text-stone-400 italic">Selecciona al menos una columna numérica para graficar.</div>;

    const nombreY = nombresEjesY[0];
    const unicaSerie = nombresEjesY.length === 1;
    const paletaActiva = obtenerPaletaActivaBloque(bloque);
    const altoContenedor = Math.max(200, Math.min(900, Number(bloque.alturaGrafica) || 380));
    const radioPastel = Math.max(80, Math.min(220, Math.round(altoContenedor * 0.34)));

    // Ordenador para que el cuadro flotante (Tooltip) al pasar el mouse respete el orden #1, #2, #3... y NO el orden alfabético
    const ordenarTooltipSegunSeleccion = (item: any) => {
      const clave = String(item?.dataKey ?? item?.name ?? '');
      const pos = nombresEjesY.indexOf(clave);
      return pos === -1 ? 999 : pos;
    };

    // Cálculo seguro del dominio numérico (evita bloqueos si min >= max de los nuevos datos pegados)
    let maxDatoReal = 0;
    datos.forEach((d: any) => {
      nombresEjesY.forEach(n => {
        const v = Number(d[n]) || 0;
        if (v > maxDatoReal) maxDatoReal = v;
      });
    });

    let minValConfig = bloque.ejeYMin !== undefined && String(bloque.ejeYMin).trim() !== '' && isFinite(Number(bloque.ejeYMin))
      ? Number(bloque.ejeYMin)
      : 0;
    let maxValConfig: number | 'auto' = bloque.ejeYMax !== undefined && String(bloque.ejeYMax).trim() !== '' && isFinite(Number(bloque.ejeYMax))
      ? Number(bloque.ejeYMax)
      : 'auto';

    if (typeof maxValConfig === 'number' && minValConfig >= maxValConfig) {
      minValConfig = 0;
      maxValConfig = 'auto';
    } else if (maxValConfig === 'auto' && maxDatoReal > 0 && minValConfig >= maxDatoReal) {
      minValConfig = 0;
    }

    const dominioNumerico: [number, number | 'auto'] = [minValConfig, maxValConfig];

    const barSizeProp = bloque.grosorBarra && Number(bloque.grosorBarra) > 0 ? Number(bloque.grosorBarra) : undefined;
    const categoryGapProp = bloque.espacioEntreBarras !== undefined ? `${bloque.espacioEntreBarras}%` : '20%';

    const formatearEtiqueta = (v: any, nombreSerie?: string) => {
      const num = Number(v);
      if (isNaN(num)) return String(v);
      const llevaPorc = nombreSerie ? (columnasPorcentaje[nombreSerie] || sufijoPorcentaje) : sufijoPorcentaje;
      return `${num.toLocaleString('es-CO', { maximumFractionDigits: 2 })}${llevaPorc ? '%' : ''}`;
    };

    switch (bloque.tipoGrafica) {
      case 'barras':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <BarChart data={datos} barCategoryGap={categoryGapProp} barGap={2} margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" hide={!!bloque.ocultarEjeX} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis domain={dominioNumerico} allowDataOverflow={false} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip itemSorter={ordenarTooltipSegunSeleccion} formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), String(name).replace(/․/g, '.')]} cursor={{ fill: '#f5f5f4' }} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              {!bloque.ocultarLeyenda && <Legend content={(props) => renderizarLeyendaEnOrden(props, nombresEjesY, paletaActiva, bloque)} />}
              {nombresEjesY.map((nombre, i) => (
                <Bar key={nombre} dataKey={nombre} barSize={barSizeProp} fill={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} radius={[6, 6, 0, 0]} isAnimationActive={false}>
                  {unicaSerie && datos.map((_: any, idx: number) => (
                    <Cell key={`cell-bar-${idx}`} fill={bloque.coloresCustom?.[idx] || paletaActiva[idx % paletaActiva.length]} />
                  ))}
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                </Bar>
              ))}
            </BarChart>
          </ResponsiveContainer>
        );

      case 'barras_horizontales':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <BarChart layout="vertical" data={datos} barCategoryGap={categoryGapProp} barGap={2} margin={{ top: 20, right: 55, left: bloque.ocultarEjeX ? 10 : 40, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e7e5e4" />
              <XAxis type="number" domain={dominioNumerico} allowDataOverflow={false} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <YAxis type="category" dataKey="name" hide={!!bloque.ocultarEjeX} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} width={90} />
              <RechartsTooltip itemSorter={ordenarTooltipSegunSeleccion} formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), String(name).replace(/․/g, '.')]} cursor={{ fill: '#f5f5f4' }} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              {!bloque.ocultarLeyenda && <Legend content={(props) => renderizarLeyendaEnOrden(props, nombresEjesY, paletaActiva, bloque)} />}
              {nombresEjesY.map((nombre, i) => (
                <Bar key={nombre} dataKey={nombre} barSize={barSizeProp} fill={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} radius={[0, 6, 6, 0]} isAnimationActive={false}>
                  {unicaSerie && datos.map((_: any, idx: number) => (
                    <Cell key={`cell-hbar-${idx}`} fill={bloque.coloresCustom?.[idx] || paletaActiva[idx % paletaActiva.length]} />
                  ))}
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="right" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                </Bar>
              ))}
            </BarChart>
          </ResponsiveContainer>
        );

      case 'barras_apiladas':
      case 'barras_3d_apiladas':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <BarChart data={datos} barCategoryGap={categoryGapProp} margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <defs>
                <filter id={`shadowStack-${bloque.id}`} x="-20%" y="-20%" width="140%" height="140%">
                  <feDropShadow dx="3" dy="5" stdDeviation="3" floodColor="#000000" floodOpacity="0.22" />
                </filter>
              </defs>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" hide={!!bloque.ocultarEjeX} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis domain={dominioNumerico} allowDataOverflow={false} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip itemSorter={ordenarTooltipSegunSeleccion} formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), String(name).replace(/․/g, '.')]} cursor={{ fill: '#f5f5f4' }} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              {!bloque.ocultarLeyenda && <Legend content={(props) => renderizarLeyendaEnOrden(props, nombresEjesY, paletaActiva, bloque)} />}
              {nombresEjesY.map((nombre, i) => (
                <Bar 
                  key={nombre} 
                  dataKey={nombre} 
                  stackId="a" 
                  barSize={barSizeProp}
                  fill={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} 
                  radius={i === nombresEjesY.length - 1 ? [8, 8, 0, 0] : [0, 0, 0, 0]} 
                  style={bloque.tipoGrafica === 'barras_3d_apiladas' ? { filter: `url(#shadowStack-${bloque.id})` } : undefined}
                  isAnimationActive={false}
                >
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="inside" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#ffffff', fontSize: 11, fontWeight: 'bold' }} />}
                </Bar>
              ))}
            </BarChart>
          </ResponsiveContainer>
        );

      case 'barras_4d':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <BarChart data={datos} barCategoryGap={categoryGapProp} barGap={2} margin={{ top: 30, right: 30, left: 0, bottom: 5 }}>
              <defs>
                <filter id={`shadow4d-${bloque.id}`} x="-20%" y="-20%" width="140%" height="140%">
                  <feDropShadow dx="4" dy="6" stdDeviation="4" floodColor="#000000" floodOpacity="0.25" />
                </filter>
              </defs>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" hide={!!bloque.ocultarEjeX} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12, fontWeight: 'bold' }} />
              <YAxis domain={dominioNumerico} allowDataOverflow={false} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip itemSorter={ordenarTooltipSegunSeleccion} formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), String(name).replace(/․/g, '.')]} cursor={{ fill: '#f5f5f4' }} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.15)' }} />
              {!bloque.ocultarLeyenda && <Legend content={(props) => renderizarLeyendaEnOrden(props, nombresEjesY, paletaActiva, bloque)} />}
              {nombresEjesY.map((nombre, i) => {
                const baseColor = bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length];
                return (
                  <Bar 
                    key={nombre} 
                    dataKey={nombre} 
                    barSize={barSizeProp}
                    fill={baseColor} 
                    radius={[12, 12, 4, 4]} 
                    style={{ filter: `url(#shadow4d-${bloque.id})` }}
                    isAnimationActive={false}
                  >
                    {unicaSerie && datos.map((_: any, idx: number) => (
                      <Cell key={`cell-4dbar-${idx}`} fill={bloque.coloresCustom?.[idx] || paletaActiva[idx % paletaActiva.length]} />
                    ))}
                    {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#1c1917', fontSize: 12, fontWeight: 900 }} />}
                  </Bar>
                );
              })}
            </BarChart>
          </ResponsiveContainer>
        );
      
      case 'torta':
      case 'torta_rellena':
      case 'dona':
      case 'semicirculo':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <PieChart>
              <Pie 
                data={datos} 
                cx="50%" 
                cy={bloque.tipoGrafica === 'semicirculo' ? '72%' : '50%'} 
                startAngle={bloque.tipoGrafica === 'semicirculo' ? 180 : 0}
                endAngle={bloque.tipoGrafica === 'semicirculo' ? 0 : 360}
                innerRadius={bloque.tipoGrafica === 'dona' ? Math.round(radioPastel * 0.62) : (bloque.tipoGrafica === 'semicirculo' ? Math.round(radioPastel * 0.5) : 0)} 
                outerRadius={radioPastel} 
                paddingAngle={bloque.tipoGrafica === 'dona' || bloque.tipoGrafica === 'semicirculo' ? 3 : (bloque.tipoGrafica === 'torta_rellena' ? 0 : 1)} 
                dataKey={nombreY} 
                nameKey="name" 
                label={bloque.mostrarValores ? ({ name, value }) => `${bloque.ocultarEjeX ? '' : `${name}: `}${formatearEtiqueta(value, nombreY)}` : false}
                isAnimationActive={false}
              >
                {datos.map((_, i) => (
                  <Cell 
                    key={`cell-${i}`} 
                    fill={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} 
                    stroke={bloque.tipoGrafica === 'torta_rellena' ? '#ffffff' : 'none'}
                    strokeWidth={bloque.tipoGrafica === 'torta_rellena' ? 1.5 : 0}
                  />
                ))}
              </Pie>
              <RechartsTooltip formatter={(value: any) => formatearEtiqueta(value, nombreY)} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              {!bloque.ocultarLeyenda && !bloque.ocultarEjeX && <Legend verticalAlign="bottom" height={36} />}
            </PieChart>
          </ResponsiveContainer>
        );

      case 'rosa_polar': {
        const maxValRosa = Math.max(...datos.map((d: any) => Number(d[nombreY]) || 0), 1);
        const datosIguales = datos.map((d: any) => ({ ...d, _anguloIgual: 1, _valorReal: Number(d[nombreY]) || 0 }));
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <PieChart>
              <Pie
                data={datosIguales}
                cx="50%"
                cy="50%"
                innerRadius={0}
                outerRadius={radioPastel}
                dataKey="_anguloIgual"
                nameKey="name"
                isAnimationActive={false}
                label={bloque.mostrarValores ? ({ payload }: any) => `${bloque.ocultarEjeX ? '' : `${payload.name}: `}${formatearEtiqueta(payload._valorReal, nombreY)}` : false}
                shape={(props: any) => {
                  const { cx, cy, startAngle, endAngle, fill, payload } = props;
                  const val = Number(payload._valorReal) || 0;
                  const r = 22 + (val / maxValRosa) * (radioPastel - 22);
                  return (
                    <Sector
                      cx={cx}
                      cy={cy}
                      innerRadius={0}
                      outerRadius={r}
                      startAngle={startAngle}
                      endAngle={endAngle}
                      fill={fill}
                      stroke="#ffffff"
                      strokeWidth={2}
                    />
                  );
                }}
              >
                {datosIguales.map((_, i) => (
                  <Cell key={`cell-rosa-${i}`} fill={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} />
                ))}
              </Pie>
              <RechartsTooltip formatter={(_: any, __: any, item: any) => [formatearEtiqueta(item?.payload?._valorReal, nombreY), item?.payload?.name]} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              {!bloque.ocultarLeyenda && !bloque.ocultarEjeX && <Legend verticalAlign="bottom" height={36} />}
            </PieChart>
          </ResponsiveContainer>
        );
      }

      case 'torta_3d_rellena':
      case 'torta_4d': {
        const sinHueco = bloque.tipoGrafica === 'torta_3d_rellena';
        return (
          <div className="relative w-full flex flex-col items-center justify-center" style={{ height: altoContenedor }}>
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <defs>
                  <filter id={`pie4d-${bloque.id}`} x="-30%" y="-30%" width="160%" height="160%">
                    <feDropShadow dx="0" dy="14" stdDeviation="8" floodColor="#0f172a" floodOpacity="0.35" />
                  </filter>
                </defs>
                <Pie
                  data={datos}
                  cx="50%"
                  cy="55%"
                  innerRadius={sinHueco ? 0 : Math.round(radioPastel * 0.44)}
                  outerRadius={radioPastel}
                  paddingAngle={sinHueco ? 0 : 4}
                  dataKey={nombreY}
                  isAnimationActive={false}
                  legendType="none"
                  tooltipType="none"
                >
                  {datos.map((_, i) => (
                    <Cell 
                      key={`depth-${i}`} 
                      fill={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} 
                      style={{ filter: 'brightness(0.62)' }}
                      stroke="none"
                    />
                  ))}
                </Pie>
                <Pie 
                  data={datos} 
                  cx="50%" 
                  cy="47%" 
                  innerRadius={sinHueco ? 0 : Math.round(radioPastel * 0.44)} 
                  outerRadius={radioPastel} 
                  paddingAngle={sinHueco ? 0 : 4} 
                  dataKey={nombreY} 
                  nameKey="name" 
                  stroke="#ffffff"
                  strokeWidth={2}
                  style={{ filter: `url(#pie4d-${bloque.id})` }}
                  label={bloque.mostrarValores ? ({ name, value }) => `${bloque.ocultarEjeX ? '' : `${name} `}(${formatearEtiqueta(value, nombreY)})` : false}
                  isAnimationActive={false}
                >
                  {datos.map((_, i) => (
                    <Cell 
                      key={`cell-4d-${i}`} 
                      fill={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} 
                    />
                  ))}
                </Pie>
                <RechartsTooltip formatter={(value: any) => formatearEtiqueta(value, nombreY)} contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.2)' }} />
                {!bloque.ocultarLeyenda && !bloque.ocultarEjeX && <Legend verticalAlign="bottom" height={36} />}
              </PieChart>
            </ResponsiveContainer>
          </div>
        );
      }

      case 'barras_radiales':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <RadialBarChart cx="50%" cy="50%" innerRadius="20%" outerRadius="95%" barSize={barSizeProp || 22} data={datos}>
              <RadialBar background dataKey={nombreY} cornerRadius={10} isAnimationActive={false}>
                {bloque.mostrarValores && <LabelList position="insideStart" fill="#1e293b" fontSize={11} fontWeight="bold" formatter={(v: any) => formatearEtiqueta(v, nombreY)} />}
              </RadialBar>
              <RechartsTooltip formatter={(value: any) => formatearEtiqueta(value, nombreY)} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              {!bloque.ocultarLeyenda && !bloque.ocultarEjeX && <Legend iconSize={10} layout="vertical" verticalAlign="middle" wrapperStyle={{ right: 0 }} />}
            </RadialBarChart>
          </ResponsiveContainer>
        );

      case 'embudo':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <FunnelChart>
              <RechartsTooltip formatter={(value: any) => formatearEtiqueta(value, nombreY)} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Funnel dataKey={nombreY} data={datos} isAnimationActive={false}>
                {!bloque.ocultarEjeX && <LabelList position="right" fill="#292524" stroke="none" dataKey="name" style={{ fontSize: 12, fontWeight: 'bold' }} />}
                {bloque.mostrarValores && <LabelList position="center" fill="#1e293b" stroke="none" dataKey={nombreY} formatter={(v: any) => formatearEtiqueta(v, nombreY)} style={{ fontSize: 12, fontWeight: 900 }} />}
              </Funnel>
            </FunnelChart>
          </ResponsiveContainer>
        );

      case 'treemap':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <Treemap
              data={datos}
              dataKey={nombreY}
              nameKey="name"
              aspectRatio={4 / 3}
              stroke="#ffffff"
              isAnimationActive={false}
              fill={edicionActiva?.color_principal || '#f97316'}
            >
              <RechartsTooltip formatter={(value: any) => formatearEtiqueta(value, nombreY)} />
            </Treemap>
          </ResponsiveContainer>
        );

      case 'dispersion':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <ScatterChart margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e7e5e4" />
              <XAxis type="category" dataKey="name" name="Categoría" hide={!!bloque.ocultarEjeX} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis type="number" dataKey={nombreY} name={String(nombreY).replace(/․/g, '.')} domain={dominioNumerico} allowDataOverflow={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip itemSorter={ordenarTooltipSegunSeleccion} cursor={{ strokeDasharray: '3 3' }} formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), String(name).replace(/․/g, '.')]} />
              {!bloque.ocultarLeyenda && <Legend content={(props) => renderizarLeyendaEnOrden(props, nombresEjesY, paletaActiva, bloque)} />}
              {nombresEjesY.map((nombre, i) => (
                <Scatter key={nombre} name={nombre} data={datos} isAnimationActive={false} fill={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]}>
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                </Scatter>
              ))}
            </ScatterChart>
          </ResponsiveContainer>
        );

      case 'lineas':
      case 'lineas_curvas':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <LineChart data={datos} margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" hide={!!bloque.ocultarEjeX} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis domain={dominioNumerico} allowDataOverflow={false} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip itemSorter={ordenarTooltipSegunSeleccion} formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), String(name).replace(/․/g, '.')]} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              {!bloque.ocultarLeyenda && <Legend content={(props) => renderizarLeyendaEnOrden(props, nombresEjesY, paletaActiva, bloque)} />}
              {nombresEjesY.map((nombre, i) => (
                <Line type={bloque.tipoGrafica === 'lineas_curvas' ? 'natural' : 'linear'} key={nombre} dataKey={nombre} stroke={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} strokeWidth={3.5} dot={{ r: 4, strokeWidth: 2 }} activeDot={{ r: 6 }} isAnimationActive={false}>
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                </Line>
              ))}
            </LineChart>
          </ResponsiveContainer>
        );

      case 'area':
      case 'area_apilada':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <AreaChart data={datos} margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" hide={!!bloque.ocultarEjeX} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis domain={dominioNumerico} allowDataOverflow={false} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip itemSorter={ordenarTooltipSegunSeleccion} formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), String(name).replace(/․/g, '.')]} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              {!bloque.ocultarLeyenda && <Legend content={(props) => renderizarLeyendaEnOrden(props, nombresEjesY, paletaActiva, bloque)} />}
              {nombresEjesY.map((nombre, i) => (
                <Area type="monotone" key={nombre} dataKey={nombre} stackId={bloque.tipoGrafica === 'area_apilada' ? '1' : undefined} stroke={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} fill={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} fillOpacity={0.65} isAnimationActive={false}>
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                </Area>
              ))}
            </AreaChart>
          </ResponsiveContainer>
        );

      case 'combinada':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <ComposedChart data={datos} barCategoryGap={categoryGapProp} margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" hide={!!bloque.ocultarEjeX} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis domain={dominioNumerico} allowDataOverflow={false} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip itemSorter={ordenarTooltipSegunSeleccion} formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), String(name).replace(/․/g, '.')]} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              {!bloque.ocultarLeyenda && <Legend content={(props) => renderizarLeyendaEnOrden(props, nombresEjesY, paletaActiva, bloque)} />}
              {nombresEjesY.map((nombre, i) => {
                const color = bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length];
                if (i === nombresEjesY.length - 1 && nombresEjesY.length > 1) {
                  return (
                    <Line key={nombre} type="monotone" dataKey={nombre} stroke={color} strokeWidth={3} dot={{ r: 4 }} isAnimationActive={false}>
                      {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                    </Line>
                  );
                }
                return (
                  <Bar key={nombre} dataKey={nombre} barSize={barSizeProp} fill={color} radius={[6, 6, 0, 0]} isAnimationActive={false}>
                    {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                  </Bar>
                );
              })}
            </ComposedChart>
          </ResponsiveContainer>
        );

      case 'radar':
      case 'radar_relleno':
        return (
          <ResponsiveContainer width="100%" height={altoContenedor}>
            <RadarChart cx="50%" cy="50%" outerRadius="80%" data={datos}>
              <PolarGrid gridType={bloque.tipoGrafica === 'radar_relleno' ? 'circle' : 'polygon'} stroke="#e7e5e4" />
              {!bloque.ocultarEjeX && <PolarAngleAxis dataKey="name" tick={{ fill: '#78716c', fontSize: 12 }} />}
              <PolarRadiusAxis angle={30} domain={dominioNumerico} tick={{ fill: '#a8a29e', fontSize: 10 }} />
              <RechartsTooltip itemSorter={ordenarTooltipSegunSeleccion} formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), String(name).replace(/․/g, '.')]} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              {!bloque.ocultarLeyenda && <Legend content={(props) => renderizarLeyendaEnOrden(props, nombresEjesY, paletaActiva, bloque)} />}
              {nombresEjesY.map((nombre, i) => (
                <Radar key={nombre} name={nombre} dataKey={nombre} stroke={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} strokeWidth={2.5} fill={bloque.coloresCustom?.[i] || paletaActiva[i % paletaActiva.length]} fillOpacity={bloque.tipoGrafica === 'radar_relleno' ? 0.75 : 0.5} isAnimationActive={false}>
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 10, fontWeight: 'bold' }} />}
                </Radar>
              ))}
            </RadarChart>
          </ResponsiveContainer>
        );

      case 'mapa_colombia':
        const valorMaximo = Math.max(...datos.map((d: any) => Number(d.valorMapa) || 0), 1);
        const colorMapa = bloque.coloresCustom?.[0] || edicionActiva.color_principal || "#f97316";
        
        const escalaColor = scaleLinear<string>()
          .domain([0, valorMaximo])
          .range(["#f1f5f9", colorMapa]); 

        const estiloMapa: any = {
          default: { outline: "none" },
          hover: { fill: "#fde047", stroke: "#0f172a", strokeWidth: 1.2, outline: "none", cursor: "pointer", transition: "all 150ms" },
          pressed: { fill: "#facc15", outline: "none" }
        };

        return (
          <div className="w-full relative flex flex-col md:flex-row items-center rounded-2xl border border-stone-200 overflow-hidden bg-transparent" style={{ minHeight: `${altoContenedor}px` }}>
            {tooltipMapa && (
              <div 
                className="fixed z-9999 pointer-events-none bg-stone-900/95 text-white px-4 py-2.5 rounded-xl shadow-2xl border border-stone-700 text-xs font-sans transform -translate-x-1/2 -translate-y-full -mt-3 transition-opacity duration-75"
                style={{ left: tooltipMapa.x, top: tooltipMapa.y }}
              >
                <span className="block text-[10px] uppercase tracking-widest text-stone-400 font-bold mb-0.5">{tooltipMapa.nombre}</span>
                <span className="block text-sm font-black text-amber-300">{(nombreY || 'Total').replace(/․/g, '.')}: {tooltipMapa.valor}</span>
              </div>
            )}

            <div className="w-full" style={{ height: `${altoContenedor}px` }}>
              <ComposableMap width={800} height={600} projection="geoMercator" projectionConfig={{ scale: 1800, center: [-74, 4.5] }} className="w-full h-full bg-transparent">
                <ZoomableGroup center={[-74, 4.5]} zoom={1} maxZoom={5}>
                  <Geographies geography={colombiaGeoJson || "/colombia.json"}>
                    {({ geographies }) =>
                      geographies.map((geo) => {
                        const nombreOriginal = geo.properties?.name || "";
                        const nombreLimpio = normalizarTexto(nombreOriginal);
                        let datoEncontrado = datos.find((d: any) => normalizarTexto(d.name) === nombreLimpio);
                        if (!datoEncontrado) {
                          datoEncontrado = datos.find((d: any) => {
                            const userT = normalizarTexto(d.name);
                            return userT && (nombreLimpio.includes(userT) || userT.includes(nombreLimpio));
                          });
                        }
                        const valor = datoEncontrado ? (Number(datoEncontrado.valorMapa) || 0) : 0;
                        const colorRelleno = datoEncontrado && valor > 0 ? escalaColor(valor) : "#e2e8f0";

                        return (
                          <Geography
                            key={geo.rsmKey}
                            geography={geo}
                            fill={colorRelleno}
                            stroke="#ffffff"
                            strokeWidth={0.5}
                            onMouseEnter={(e: React.MouseEvent) => setTooltipMapa({ nombre: nombreOriginal, valor: formatearEtiqueta(valor, nombreY), x: e.clientX, y: e.clientY })}
                            onMouseMove={(e: React.MouseEvent) => setTooltipMapa(prev => prev ? { ...prev, x: e.clientX, y: e.clientY } : null)}
                            onMouseLeave={() => setTooltipMapa(null)}
                            style={estiloMapa}
                          />
                        );
                      })
                    }
                  </Geographies>
                </ZoomableGroup>
              </ComposableMap>
            </div>

            {bloque.mostrarValores && (
              <div className="w-full md:w-72 max-h-105 overflow-y-auto bg-white/90 backdrop-blur border-t md:border-t-0 md:border-l border-stone-200 p-4 font-sans text-xs space-y-2 shrink-0">
                <p className="font-black uppercase tracking-widest text-stone-400 text-[10px] mb-2">Valores por Región</p>
                {datos.map((d: any, idx: number) => (
                  <div key={idx} className="flex justify-between items-center py-1.5 border-b border-stone-100 last:border-b-0">
                    <span className="font-semibold text-stone-700 truncate pr-2">{d.name}</span>
                    <span className="font-black text-blue-700 bg-blue-50 px-2 py-0.5 rounded">{formatearEtiqueta(d.valorMapa, nombreY)}</span>
                  </div>
                ))}
              </div>
            )}
          </div>
        );

      default:
        return null;
    }
  };

  const renderizarBotonesEnTexto = (html: string) => {
    if (!html) return '';

    const htmlSinNbsp = html.replace(/&nbsp;|&#160;|\u00A0/gi, ' ');
    const htmlConSaltosPreservados = htmlSinNbsp
      .replace(/<p(\s+[^>]*)?>\s*<br\s*\/?>\s*<\/p>/gi, '<p$1>&nbsp;</p>')
      .replace(/<p(\s+[^>]*)?>\s*<\/p>/gi, '<p$1>&nbsp;</p>');

    const htmlConLinksForzados = htmlConSaltosPreservados.replace(/<a\b([^>]*data-custom-link="true"[^>]*)>/gi, (_, attrs) => {
      const boldMatch = attrs.match(/data-bold="([^"]*)"/i);
      const underMatch = attrs.match(/data-underline="([^"]*)"/i);
      const strikeMatch = attrs.match(/data-strike="([^"]*)"/i);
      const italicMatch = attrs.match(/data-italic="([^"]*)"/i);
      const colorMatch = attrs.match(/data-color="([^"]*)"/i);

      const esBold = boldMatch ? boldMatch[1] === 'true' : true;
      const esUnder = underMatch ? underMatch[1] === 'true' : true;
      const esStrike = strikeMatch ? strikeMatch[1] === 'true' : false;
      const esItalic = italicMatch ? italicMatch[1] === 'true' : false;
      const colorHex = colorMatch ? colorMatch[1] : '#1d4ed8';

      const decos: string[] = [];
      if (esUnder) decos.push('underline');
      if (esStrike) decos.push('line-through');
      const decoCss = decos.length > 0 ? decos.join(' ') : 'none';
      const weightCss = esBold ? '700' : '400';
      const italicCss = esItalic ? 'italic' : 'normal';

      const styleForzado = `color: ${colorHex} !important; font-weight: ${weightCss} !important; font-style: ${italicCss} !important; text-decoration: ${decoCss} !important; cursor: pointer;`;
      const attrsSinStyle = attrs.replace(/\s*style="[^"]*"/gi, '');
      return `<a ${attrsSinStyle} style="${styleForzado}">`;
    });

    return htmlConLinksForzados.replace(/\[\[BOTON\|([\s\S]*?)\]\]/g, (_, contenidoInterno) => {
      const limpio = String(contenidoInterno)
        .replace(/<[^>]*>/g, '')
        .replace(/&nbsp;/g, ' ')
        .trim();
      const partes = limpio.split('|').map(p => p.trim());

      const texto = partes[0] || 'Ver más';
      const url = partes[1] || '#';

      let estiloRaw = 'solido';
      let cFondo = '#1d4ed8';
      let cTexto = '#ffffff';
      let cGlow = '';
      let tamanoRaw = 'mediano';
      let fuenteRaw = 'arial';

      if (partes.length === 3 && partes[2].startsWith('#')) {
        cFondo = partes[2];
      } else {
        estiloRaw = (partes[2] || 'solido')
          .toLowerCase()
          .normalize("NFD")
          .replace(/[\u0300-\u036f]/g, "");
        cFondo = partes[3] || '#1d4ed8';
        cTexto = partes[4] || '#ffffff';
        cGlow = partes[5] || cFondo;
        tamanoRaw = (partes[6] || 'mediano').toLowerCase();
        fuenteRaw = partes[7] || 'arial';
      }

      let padBtn = '9px 22px';
      let fontBtn = '12px';
      if (tamanoRaw === 'pequeno') { padBtn = '6px 14px'; fontBtn = '10px'; }
      else if (tamanoRaw === 'grande') { padBtn = '13px 30px'; fontBtn = '14px'; }
      else if (tamanoRaw === 'extragrande') { padBtn = '16px 38px'; fontBtn = '16px'; }

      const fontFamilyCss = NOMBRES_FUENTES_WORD[fuenteRaw]?.family || 'sans-serif';

      let extraStyle = `background-color: ${cFondo} !important; color: ${cTexto} !important; border: 2px solid ${cFondo} !important; border-radius: 12px !important; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.15) !important;`;

      if (estiloRaw.includes('pildora') || estiloRaw.includes('oval')) {
        extraStyle = `background-color: ${cFondo} !important; color: ${cTexto} !important; border: 2px solid ${cFondo} !important; border-radius: 9999px !important; box-shadow: 0 6px 14px -2px rgba(0,0,0,0.18) !important;`;
      } else if (estiloRaw.includes('contorno')) {
        extraStyle = `background-color: transparent !important; color: ${cFondo} !important; border: 2px solid ${cFondo} !important; border-radius: 12px !important; box-shadow: none !important;`;
      } else if (estiloRaw.includes('3d')) {
        extraStyle = `background-color: ${cFondo} !important; color: ${cTexto} !important; border: 1px solid rgba(255,255,255,0.25) !important; border-bottom: 5px solid rgba(0,0,0,0.35) !important; border-radius: 12px !important; box-shadow: 0 6px 0 0 ${cFondo}cc, 0 10px 15px -3px rgba(0,0,0,0.28) !important; transform: translateY(-2px) !important;`;
      } else if (estiloRaw.includes('4d')) {
        const luz = cGlow || cFondo;
        extraStyle = `background-color: ${cFondo} !important; color: ${cTexto} !important; border: 1.5px solid rgba(255,255,255,0.65) !important; border-radius: 9999px !important; box-shadow: 0 0 24px 5px ${luz}, inset 0 0 12px rgba(255,255,255,0.45) !important;`;
      }

      const esLinkInterno = url.startsWith('#pagina:') || (!url.startsWith('http') && url.startsWith('#'));
      const targetAttr = esLinkInterno ? `data-internal-page="${url.replace(/^#pagina:/i, '').replace(/^#/, '')}"` : 'target="_blank" rel="noreferrer"';

      return `<a href="${url}" ${targetAttr} data-inline-btn="true" style="${extraStyle} padding: ${padBtn} !important; font-family: ${fontFamilyCss} !important; font-weight: 900 !important; font-size: ${fontBtn} !important; line-height: 1.2 !important; text-transform: uppercase !important; letter-spacing: 0.06em !important; text-decoration: none !important; display: inline-block !important; margin: 6px 6px !important; vertical-align: middle !important; cursor: pointer !important; transition: transform 0.2s ease, filter 0.2s ease !important;">${texto}</a>`;
    });
  };

  const seccionActiva = secciones[paginaActual] || { tipo: 'portada' };
  const colorFondoApp = seccionActiva.tipo === 'portada'
    ? (configuracion.colorFondoPortada || '#fdfbf7')
    : (seccionActiva.pagina.colorFondo || configuracion.colorFondoPaginasGlobal || '#fdfbf7');

  const mostrarMarcaAgua = () => {
    if (!configuracion.marcaAguaUrl) return false;
    if (seccionActiva.tipo === 'portada') return false;
    if (configuracion.modoMarcaAgua === 'todas') return true;
    return (configuracion.paginasMarcaAgua || []).includes(seccionActiva.pagina.id);
  };

  const textoVacioQuill = (html?: string) => {
    if (!html) return true;
    const limpio = html.replace(/<[^>]*>/g, '').trim();
    return limpio === '' && !html.includes('<img');
  };

  const variantesVirtualizadas = {
    enter: (dir: number) => ({ opacity: 0, x: dir > 0 ? 36 : -36 }),
    center: { opacity: 1, x: 0 },
    exit: (dir: number) => ({ opacity: 0, x: dir > 0 ? -36 : 36 })
  };

  return (
    <div 
      className={`${generandoPdfPublico ? 'min-h-screen w-full overflow-visible block' : 'h-screen w-screen overflow-hidden relative'} text-stone-900 font-serif transition-colors duration-300`} 
      style={{ backgroundColor: colorFondoApp }}
      onTouchStart={onTouchStart}
      onTouchMove={onTouchMove}
      onTouchEnd={onTouchEnd}
      onClick={manejarClicContenidoInterno}
    >
      <style dangerouslySetInnerHTML={{ __html: ESTILOS_FUENTES_PUBLICO }} />

      {/* MODO DE EXPORTACIÓN E IMPRESIÓN DEL LIBRO COMPLETO EN PDF */}
      {generandoPdfPublico && (
        <div className="w-full bg-[#fdfbf7] font-serif text-stone-900 print-book-wrapper">
          <div className="max-w-5xl mx-auto">
            {secciones.map((sec, sIdx) => {
              if (sec.tipo === 'portada') {
                return (
                  <div 
                    key={`pdf-portada-${sIdx}`} 
                    className="print-page-sheet flex flex-col items-center justify-center text-center px-10 py-24 relative" 
                    style={{ backgroundColor: configuracion.colorFondoPortada || '#fdfbf7' }}
                  >
                    {configuracion.imagenFondoPortada && (
                      <div className="absolute inset-0 z-0 overflow-hidden">
                        <img 
                          src={configuracion.imagenFondoPortada} 
                          alt="Fondo Portada" 
                          className="w-full h-full object-cover" 
                          style={{ 
                            filter: `brightness(${configuracion.portadaBgBrightness ?? 100}%) contrast(${configuracion.portadaBgContrast ?? 100}%) blur(${configuracion.portadaBgBlur ?? 2}px)` 
                          }}
                        />
                        <div 
                          className="absolute inset-0" 
                          style={{ 
                            backgroundColor: configuracion.colorFondoPortada || '#fdfbf7',
                            opacity: (configuracion.portadaBgOverlayOpacity ?? 85) / 100
                          }} 
                        />
                      </div>
                    )}
                    <div className="relative z-10 flex flex-col items-center w-full">
                      {configuracion.logoUrl && (
                        <img src={configuracion.logoUrl} alt="Logo" style={{ height: `${configuracion.logoSize ?? 56}px` }} className="w-auto object-contain mb-8" />
                      )}
                      <div className="w-16 h-1 mb-10" style={{ backgroundColor: edicionActiva.color_principal || '#1d4ed8' }}></div>
                      {configuracion.portadaContenido ? (
                        <div 
                          className="w-full prose prose-stone max-w-none [&_h1]:text-5xl [&_h1]:font-normal [&_h1]:mb-6 [&_p]:text-xl [&_p]:italic [&_p]:text-stone-600"
                          dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(configuracion.portadaContenido) }}
                        />
                      ) : (
                        <>
                          <h1 className="text-5xl font-normal mb-6">{edicionActiva.titulo}</h1>
                          <p className="text-xl italic text-stone-500 mb-12">{configuracion.subtitulo}</p>
                        </>
                      )}
                      <div className="mt-16 flex flex-col items-center">
                        <span className="text-xs font-sans font-bold tracking-[0.3em] uppercase text-stone-400">ACOFI / Edición Oficial {edicionActiva.anio}</span>
                      </div>
                    </div>
                  </div>
                );
              }

              const p = sec.pagina;
              const tieneT = p.titulo.trim() !== '';
              const tieneSub = (p.subtitulo || '').trim() !== '';
              const mostrarNum = p.mostrarNumero !== false && tieneT;

              const cabeceraMostrarPdf = !textoVacioQuill(p.cabeceraHtml)
                ? p.cabeceraHtml
                : (configuracion.mostrarCabeceraGlobal && !textoVacioQuill(configuracion.cabeceraGlobalHtml) ? configuracion.cabeceraGlobalHtml : '');

              const pieMostrarPdf = !textoVacioQuill(p.piePaginaHtml)
                ? p.piePaginaHtml
                : (configuracion.mostrarPiePaginaGlobal && !textoVacioQuill(configuracion.piePaginaGlobalHtml) ? configuracion.piePaginaGlobalHtml : '');

              return (
                <div key={`pdf-pag-${p.id}`} className="print-page-sheet flex flex-col relative" style={{ backgroundColor: p.colorFondo || configuracion.colorFondoPaginasGlobal || '#fdfbf7' }}>
                  {p.imagenFondo && (
                    <div className="absolute inset-0 z-0 overflow-hidden pointer-events-none">
                      <img src={p.imagenFondo} alt="Fondo Página" className="w-full h-full object-cover" />
                      <div 
                        className="absolute inset-0" 
                        style={{ 
                          backgroundColor: p.colorFondo || configuracion.colorFondoPaginasGlobal || '#fdfbf7',
                          opacity: (p.opacidadFondo ?? 85) / 100
                        }} 
                      />
                    </div>
                  )}
                  <div className="relative z-10 flex flex-col flex-1">
                    {cabeceraMostrarPdf && (
                      <div className="mb-4 pb-2 border-b border-stone-200/60 prose prose-sm prose-stone max-w-none opacity-80 print-header-keep" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(cabeceraMostrarPdf) }} />
                    )}
                    {(tieneT || tieneSub) && (
                      <div className="mb-6 border-b border-stone-200/60 pb-4 print-header-keep">
                        <div className="flex items-center gap-4">
                          {mostrarNum && <span className="text-4xl font-black opacity-30">{sec.numeroReal}.</span>}
                          {tieneT && <h2 className="text-3xl font-black tracking-tight">{p.titulo}</h2>}
                        </div>
                        {tieneSub && <p className="text-base text-stone-500 font-serif italic mt-2">{p.subtitulo}</p>}
                      </div>
                    )}
                    <div className="space-y-6 flex-1">
                      {p.bloques.map((b) => (
                        <div key={b.id} className={`w-full ${b.tipo === 'grafica' || b.tipo === 'firma' || b.tipo === 'imagen' ? 'print-avoid-break' : ''}`}>
                          {b.tipo === 'texto' && (
                            b.columnas === 2 || b.columnas === 3 ? (
                              <div 
                                className={`grid ${b.columnas === 2 ? 'grid-cols-2 gap-8' : 'grid-cols-3 gap-6'} p-4 rounded-2xl items-start`}
                                style={{ backgroundColor: b.colorFondo === 'transparent' ? 'transparent' : (b.colorFondo || 'transparent') }}
                              >
                                <div className="prose prose-stone max-w-none text-stone-700 leading-normal" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.contenido) }} />
                                <div className="prose prose-stone max-w-none text-stone-700 leading-normal border-l border-stone-200/70 pl-6" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.contenidoCol2 || '') }} />
                                {b.columnas === 3 && <div className="prose prose-stone max-w-none text-stone-700 leading-normal border-l border-stone-200/70 pl-6" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.contenidoCol3 || '') }} />}
                              </div>
                            ) : (
                              <div 
                                className="prose prose-stone max-w-none text-stone-700 leading-relaxed p-4 rounded-2xl"
                                style={{ backgroundColor: b.colorFondo === 'transparent' ? 'transparent' : (b.colorFondo || 'transparent') }}
                                dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.contenido) }} 
                              />
                            )
                          )}
                          {b.tipo === 'firma' && (
                            <div 
                              className={`py-4 px-4 rounded-2xl flex flex-wrap gap-12 items-start ${
                                b.alineacion === 'centro' ? 'justify-center text-center' :
                                b.alineacion === 'derecha' ? 'justify-end text-right' :
                                b.alineacion === 'espaciado' ? 'justify-around text-left' :
                                'justify-start text-left'
                              }`}
                              style={{ 
                                backgroundColor: b.colorFondo === 'transparent' ? 'transparent' : (b.colorFondo || 'transparent'),
                                paddingTop: `${Math.max(8, 20 + (b.espacioGeneralSuperior ?? 0))}px`,
                                paddingBottom: `${Math.max(8, 20 + (b.espacioGeneralInferior ?? 0))}px`
                              }}
                            >
                              {(b.firmas || []).map((f, i) => (
                                <div 
                                  key={i} 
                                  className="flex flex-col space-y-1 min-w-52"
                                  style={{ transform: `translateY(${f.espacioSuperior ?? 0}px)` }}
                                >
                                  {f.url && (
                                    <div className={`mb-2 flex ${b.alineacion === 'centro' ? 'justify-center' : b.alineacion === 'derecha' ? 'justify-end' : 'justify-start'}`}>
                                      <img 
                                        src={f.url} 
                                        alt={f.nombre} 
                                        style={{ height: `${f.tamanoFirma ?? 72}px`, maxHeight: `${f.tamanoFirma ?? 72}px`, width: 'auto' }} 
                                        className="object-contain" 
                                      />
                                    </div>
                                  )}
                                  <span className="font-sans font-black text-stone-900 text-base leading-snug">{f.nombre}</span>
                                  {f.cargo && <span className="font-sans text-stone-600 text-sm leading-snug">{f.cargo}</span>}
                                  {f.infoAdicional && <span className="font-sans text-stone-500 text-xs italic leading-snug">{f.infoAdicional}</span>}
                                </div>
                              ))}
                            </div>
                          )}
                          {b.tipo === 'boton' && (() => {
                            const szBtn = obtenerEstilosTamanoBoton(b.tamanoBoton);
                            const famBtn = NOMBRES_FUENTES_WORD[b.fuenteBoton || 'arial']?.family || 'sans-serif';
                            return (
                              <div className={`py-3 flex ${b.alineacion === 'centro' ? 'justify-center' : b.alineacion === 'derecha' ? 'justify-end' : 'justify-start'}`}>
                                <span
                                  style={{
                                    backgroundColor: b.estilo === 'contorno' ? 'transparent' : b.colorBoton,
                                    color: b.estilo === 'contorno' ? b.colorBoton : b.colorTexto,
                                    borderColor: b.colorBoton,
                                    padding: szBtn.padding,
                                    fontSize: szBtn.fontSize,
                                    fontFamily: famBtn,
                                    boxShadow: b.estilo === '4d' ? `0 0 20px 3px ${b.colorIluminacion || b.colorBoton}` : undefined
                                  }}
                                  className={`font-black uppercase tracking-wider inline-block ${b.estilo === 'pildora' || b.estilo === '4d' ? 'rounded-full' : 'rounded-xl'}`}
                                >
                                  {b.texto}
                                </span>
                              </div>
                            );
                          })()}
                          {b.tipo === 'tabla' && !b.oculto && (
                            <div className="border border-stone-300 rounded-2xl overflow-hidden my-3" style={{ backgroundColor: b.colorFondo || '#ffffff' }}>
                              <div className="bg-stone-100 p-3 font-sans font-bold text-xs uppercase print-header-keep">{b.nombre}</div>
                              <table className="w-full text-left font-sans border-collapse" style={{ fontSize: `${b.tamanoTextoTabla ?? 12}px` }}>
                                <thead>
                                  <tr className="bg-stone-200/70 border-b border-stone-300" style={{ backgroundColor: b.coloresFilas?.[0] || '#f5f5f4' }}>
                                    {b.data[0]?.map((c, ci) => <th key={ci} className="p-2 font-bold uppercase">{c}</th>)}
                                  </tr>
                                </thead>
                                <tbody className="divide-y divide-stone-200">
                                  {b.data.slice(1).map((f, fi) => (
                                    <tr key={fi} style={{ backgroundColor: b.coloresFilas?.[fi + 1] || 'transparent' }}>
                                      {f.map((cell, cii) => <td key={cii} className={`p-1.5 ${cii === 0 ? 'font-semibold' : ''}`}>{cell}</td>)}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                              {!textoVacioQuill(b.pieBloque) && <div className="p-2.5 bg-stone-50 text-xs border-t border-stone-200" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.pieBloque || '') }} />}
                            </div>
                          )}
                          {b.tipo === 'grafica' && (
                            <div className="border border-stone-200 rounded-2xl p-6 my-4 print-avoid-break" style={{ backgroundColor: b.colorFondo || '#ffffff' }}>
                              <h3 className="font-sans font-bold uppercase tracking-widest text-xs text-center mb-4">{b.nombre}</h3>
                              {renderizarGrafica(b)}
                              {!textoVacioQuill(b.pieBloque) && <div className="mt-3 pt-2 border-t border-stone-200 text-xs" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.pieBloque || '') }} />}
                            </div>
                          )}
                          {b.tipo === 'imagen' && b.url && (
                            <div className="my-4 text-center print-avoid-break">
                              <img src={b.url} alt="Imagen" className="max-h-80 mx-auto rounded-xl object-contain" />
                            </div>
                          )}
                        </div>
                      ))}
                    </div>
                    {pieMostrarPdf && (
                      <div className="mt-8 pt-3 border-t border-stone-200 text-xs opacity-75" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(pieMostrarPdf) }} />
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* LOGO SUPERIOR GLOBAL */}
      {configuracion.logoUrl && !generandoPdfPublico && (
        <div className="absolute top-6 left-6 md:top-8 md:left-12 z-40 pointer-events-none no-print">
          <img 
            src={configuracion.logoUrl} 
            alt="Logo" 
            style={{ 
              height: `${configuracion.logoSize ?? 56}px`, 
              filter: `brightness(${configuracion.logoBrightness ?? 100}%)` 
            }} 
            className="w-auto object-contain drop-shadow-sm transition-all duration-300" 
          />
        </div>
      )}

      {/* BARRA FLOTANTE SUPERIOR DERECHA */}
      {!generandoPdfPublico && (
        <div className="absolute top-5 right-5 md:top-7 md:right-10 z-40 flex flex-wrap items-center justify-end gap-2 no-print">
          
          {edicionesDisponibles.length > 1 && (
            <div className="flex items-center gap-1.5 bg-white/85 backdrop-blur-xl border border-stone-200/90 px-3 py-2 rounded-full shadow-sm font-sans text-xs font-bold text-stone-700">
              <Calendar size={13} className="text-stone-400" />
              <select
                value={edicionActiva.anio}
                onChange={(e) => router.push(`/?edicion=${e.target.value}`)}
                className="bg-transparent outline-none cursor-pointer font-black text-stone-800"
                title="Cambiar de edición anual"
              >
                {edicionesDisponibles.map(ed => (
                  <option key={ed.id} value={ed.anio}>Edición {ed.anio}</option>
                ))}
              </select>
            </div>
          )}

          <button
            type="button"
            onClick={descargarLibroCompletoPdf}
            disabled={generandoPdfPublico}
            className="flex items-center gap-2 bg-white/85 hover:bg-white backdrop-blur-xl border border-stone-200/90 hover:border-emerald-600 text-stone-700 hover:text-emerald-700 px-3.5 py-2.5 rounded-full shadow-md shadow-stone-200/60 transition-all font-sans text-xs font-bold cursor-pointer"
            title="Descargar el libro completo en PDF"
          >
            <Printer size={14} className="text-emerald-600" />
            <span className="hidden md:inline">{generandoPdfPublico ? 'Preparando PDF...' : 'Descargar PDF'}</span>
          </button>

          <button
            type="button"
            onClick={compartirPaginaActual}
            className="flex items-center gap-1.5 bg-white/85 hover:bg-white backdrop-blur-xl border border-stone-200/90 hover:border-blue-600 text-stone-700 hover:text-blue-700 px-3 py-2.5 rounded-full shadow-md shadow-stone-200/60 transition-all font-sans text-xs font-bold cursor-pointer"
            title="Copiar enlace directo a esta página"
          >
            {enlaceCompartido ? <Check size={14} className="text-emerald-600" /> : <Share2 size={14} />}
            {enlaceCompartido && <span className="text-[11px] text-emerald-700">¡Enlace copiado!</span>}
          </button>

          <button
            type="button"
            onClick={alternarPantallaCompleta}
            className="hidden sm:flex items-center justify-center bg-white/85 hover:bg-white backdrop-blur-xl border border-stone-200/90 hover:border-stone-400 text-stone-700 p-2.5 rounded-full shadow-md shadow-stone-200/60 transition-all cursor-pointer"
            title="Modo Pantalla Completa"
          >
            {esPantallaCompleta ? <Minimize2 size={14} /> : <Maximize2 size={14} />}
          </button>

          <button
            type="button"
            onClick={() => setBuscadorAbierto(true)}
            className="flex items-center gap-2.5 bg-white/90 hover:bg-white backdrop-blur-xl border border-stone-200/90 hover:border-blue-600 text-stone-700 hover:text-blue-700 px-4 py-2.5 rounded-full shadow-md shadow-stone-200/60 transition-all font-sans text-xs font-bold group cursor-pointer"
            title="Buscar ingeniería, tabla, ciudad o ir a cualquier página (Ctrl + K)"
          >
            <Search size={15} className="text-stone-400 group-hover:text-blue-600 transition-colors" />
            <span className="hidden sm:inline">Buscar en el libro...</span>
            <kbd className="hidden lg:inline-block px-2 py-0.5 text-[10px] font-black bg-stone-100 text-stone-500 rounded-md border border-stone-200">
              Ctrl K
            </kbd>
          </button>
        </div>
      )}

      {/* MARCA DE AGUA CENTRAL */}
      {mostrarMarcaAgua() && !generandoPdfPublico && (
        <div className="fixed inset-0 pointer-events-none z-0 flex items-center justify-center overflow-hidden no-print">
          <img 
            src={configuracion.marcaAguaUrl} 
            alt="Marca de agua" 
            style={{ 
              width: `${configuracion.marcaAguaSize ?? 40}%`, 
              opacity: (configuracion.marcaAguaOpacity ?? 5) / 100 
            }} 
            className="object-contain transition-all duration-300" 
          />
        </div>
      )}

      {/* MOTOR VIRTUALIZADO ULTRARRÁPIDO: SOLO MONTA EN EL DOM LA PÁGINA ACTIVA */}
      {!generandoPdfPublico && (
        <AnimatePresence mode="wait" initial={false} custom={direccionTransicion}>
          <motion.div
            key={seccionActiva.tipo === 'portada' ? 'portada-principal' : seccionActiva.pagina.id}
            custom={direccionTransicion}
            variants={variantesVirtualizadas}
            initial="enter"
            animate="center"
            exit="exit"
            transition={{ duration: 0.16, ease: 'easeOut' }}
            className="h-full w-screen no-print"
          >
            {seccionActiva.tipo === 'portada' ? (
              <div 
                className="h-full w-screen shrink-0 flex flex-col items-center justify-center text-center px-6 md:px-20 relative overflow-y-auto overflow-x-hidden custom-scrollbar"
              >
                {configuracion.imagenFondoPortada && (
                  <div className="absolute inset-0 z-0 overflow-hidden">
                    <img 
                      src={configuracion.imagenFondoPortada} 
                      alt="Fondo" 
                      className="w-full h-full object-cover transition-all" 
                      style={{ 
                        filter: `brightness(${configuracion.portadaBgBrightness ?? 100}%) contrast(${configuracion.portadaBgContrast ?? 100}%) blur(${configuracion.portadaBgBlur ?? 2}px)` 
                      }}
                    />
                    <div 
                      className="absolute inset-0 transition-all" 
                      style={{ 
                        backgroundColor: configuracion.colorFondoPortada || '#fdfbf7',
                        opacity: (configuracion.portadaBgOverlayOpacity ?? 85) / 100
                      }}
                    ></div>
                  </div>
                )}

                <div className="relative z-10 flex flex-col items-center w-full max-w-5xl mx-auto px-2">
                  <div className="w-16 h-1 mb-10 transition-colors" style={{ backgroundColor: edicionActiva.color_principal || '#1d4ed8' }}></div>
                  
                  {configuracion.portadaContenido ? (
                    <div 
                      className="text-center mx-auto w-full max-w-full wrap-break-word whitespace-normal overflow-hidden text-stone-900 prose prose-lg prose-stone [&_h1]:text-4xl [&_h1]:sm:text-5xl [&_h1]:md:text-6xl [&_h1]:lg:text-7xl [&_h1]:font-normal [&_h1]:tracking-tight [&_h1]:leading-[1.05] [&_h1]:mb-8 [&_h1]:wrap-break-word [&_p]:text-lg [&_p]:md:text-2xl [&_p]:font-serif [&_p]:mb-12 [&_p]:wrap-break-word"
                      dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(configuracion.portadaContenido) }} 
                    />
                  ) : (
                    <>
                      <h1 className="text-4xl sm:text-5xl md:text-6xl lg:text-7xl font-normal tracking-tight text-stone-900 w-full max-w-full wrap-break-word whitespace-normal leading-[1.05] mb-8">
                        {edicionActiva.titulo || 'Programas de Ingeniería en Colombia'}
                      </h1>
                      <p className="text-lg md:text-2xl font-serif italic text-stone-500 max-w-2xl wrap-break-word whitespace-normal mb-16">
                        {configuracion.subtitulo || 'Análisis oficial y directorio nacional de la educación superior en ingeniería.'}
                      </p>
                    </>
                  )}

                  <div className="flex flex-col items-center">
                    <span className="text-xs font-sans font-bold tracking-[0.3em] uppercase text-stone-400 mb-4">ACOFI / Edición Oficial {edicionActiva.anio}</span>
                    <BookOpen size={24} style={{ color: edicionActiva.color_principal || '#1d4ed8' }} />
                  </div>
                </div>
              </div>
            ) : (() => {
              const pagina = seccionActiva.pagina;
              const tieneTitulo = pagina.titulo.trim() !== '';
              const tieneSubtitulo = (pagina.subtitulo || '').trim() !== '';
              const mostrarNum = pagina.mostrarNumero !== false && tieneTitulo;

              const cabeceraMostrar = !textoVacioQuill(pagina.cabeceraHtml)
                ? pagina.cabeceraHtml
                : (configuracion.mostrarCabeceraGlobal && !textoVacioQuill(configuracion.cabeceraGlobalHtml) ? configuracion.cabeceraGlobalHtml : '');

              const pieMostrar = !textoVacioQuill(pagina.piePaginaHtml)
                ? pagina.piePaginaHtml
                : (configuracion.mostrarPiePaginaGlobal && !textoVacioQuill(configuracion.piePaginaGlobalHtml) ? configuracion.piePaginaGlobalHtml : '');

              return (
                <div className="h-full w-screen shrink-0 flex flex-col px-6 sm:px-12 md:px-24 lg:px-32 py-24 overflow-y-auto overflow-x-hidden custom-scrollbar relative z-10">
                  
                  {/* IMAGEN DE FONDO INDIVIDUAL DE LA PÁGINA */}
                  {pagina.imagenFondo && (
                    <div className="fixed inset-0 z-0 pointer-events-none">
                      <img src={pagina.imagenFondo} alt="Fondo de página" className="w-full h-full object-cover" />
                      <div 
                        className="absolute inset-0 backdrop-blur-[1px]" 
                        style={{ 
                          backgroundColor: pagina.colorFondo || configuracion.colorFondoPaginasGlobal || '#fdfbf7',
                          opacity: (pagina.opacidadFondo ?? 85) / 100
                        }}
                      />
                    </div>
                  )}

                  <div className="max-w-5xl mx-auto w-full flex-1 flex flex-col min-w-0 relative z-10">
                    
                    {cabeceraMostrar && (
                      <div 
                        className="mb-8 pb-4 border-b border-stone-200/60 prose prose-sm prose-stone max-w-none opacity-80"
                        dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(cabeceraMostrar) }}
                      />
                    )}

                    {(tieneTitulo || tieneSubtitulo) && (
                      <div className="mb-12 border-b border-stone-200/50 pb-6 relative z-10">
                        <div className="flex items-center gap-4">
                          {mostrarNum && (
                            <span className="text-5xl font-black opacity-30">{seccionActiva.numeroReal}.</span>
                          )}
                          {tieneTitulo && (
                            <h2 className="text-4xl md:text-5xl font-black tracking-tight wrap-break-word">
                              {pagina.titulo}
                            </h2>
                          )}
                        </div>
                        {tieneSubtitulo && (
                          <p className="text-lg md:text-xl text-stone-500 font-serif italic mt-3 wrap-break-word">
                            {pagina.subtitulo}
                          </p>
                        )}
                      </div>
                    )}

                    <div className="space-y-12 pb-28 relative z-10 flex-1 min-w-0">
                      {pagina.bloques.map((bloque) => (
                        <div key={bloque.id} className="w-full min-w-0">
                          
                          {bloque.tipo === 'texto' && (
                            bloque.columnas === 2 || bloque.columnas === 3 ? (
                              <div 
                                className={`grid grid-cols-1 ${bloque.columnas === 2 ? 'md:grid-cols-2 gap-8 md:gap-12' : 'md:grid-cols-3 gap-6 md:gap-8'} p-4 sm:p-6 rounded-2xl transition-colors items-start w-full min-w-0`}
                                style={{ backgroundColor: bloque.colorFondo === 'transparent' ? 'transparent' : (bloque.colorFondo || 'transparent') }}
                              >
                                <div 
                                  className="prose prose-stone max-w-full min-w-0 text-stone-700 leading-normal wrap-break-word"
                                  dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(bloque.contenido) }}
                                />
                                <div 
                                  className="prose prose-stone max-w-full min-w-0 text-stone-700 leading-normal md:border-l md:border-stone-200/70 md:pl-8 wrap-break-word"
                                  dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(bloque.contenidoCol2 || '') }}
                                />
                                {bloque.columnas === 3 && (
                                  <div 
                                    className="prose prose-stone max-w-full min-w-0 text-stone-700 leading-normal md:border-l md:border-stone-200/70 md:pl-8 wrap-break-word"
                                    dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(bloque.contenidoCol3 || '') }}
                                  />
                                )}
                              </div>
                            ) : (
                              <div 
                                className="prose prose-lg prose-stone max-w-full min-w-0 text-stone-700 leading-relaxed p-4 sm:p-6 rounded-2xl transition-colors wrap-break-word"
                                style={{ backgroundColor: bloque.colorFondo === 'transparent' ? 'transparent' : (bloque.colorFondo || 'transparent') }}
                                dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(bloque.contenido) }}
                              />
                            )
                          )}

                          {/* RENDERIZADO DEL BLOQUE DE FIRMAS CON MOVIMIENTO INDIVIDUAL SIN AFECTAR A LAS OTRAS FIRMAS */}
                          {bloque.tipo === 'firma' && (
                            <div 
                              className={`px-4 rounded-2xl flex flex-wrap gap-12 items-start ${
                                bloque.alineacion === 'centro' ? 'justify-center text-center' :
                                bloque.alineacion === 'derecha' ? 'justify-end text-right' :
                                bloque.alineacion === 'espaciado' ? 'justify-around text-left' :
                                'justify-start text-left'
                              }`}
                              style={{ 
                                backgroundColor: bloque.colorFondo === 'transparent' ? 'transparent' : (bloque.colorFondo || 'transparent'),
                                paddingTop: `${Math.max(8, 24 + (bloque.espacioGeneralSuperior ?? 0))}px`,
                                paddingBottom: `${Math.max(8, 24 + (bloque.espacioGeneralInferior ?? 0))}px`
                              }}
                            >
                              {(bloque.firmas || []).map((fItem, fIdx) => (
                                <div 
                                  key={fItem.id || fIdx} 
                                  className="flex flex-col space-y-1 min-w-56 transition-transform"
                                  style={{ transform: `translateY(${fItem.espacioSuperior ?? 0}px)` }}
                                >
                                  {fItem.url && (
                                    <div className={`mb-2 flex ${bloque.alineacion === 'centro' ? 'justify-center' : bloque.alineacion === 'derecha' ? 'justify-end' : 'justify-start'}`}>
                                      <img
                                        src={fItem.url}
                                        alt={fItem.nombre}
                                        style={{ 
                                          height: `${fItem.tamanoFirma ?? 72}px`, 
                                          maxHeight: `${fItem.tamanoFirma ?? 72}px`,
                                          width: 'auto' 
                                        }}
                                        className="object-contain"
                                      />
                                    </div>
                                  )}
                                  <span className="font-sans font-black text-stone-900 text-base leading-snug">{fItem.nombre}</span>
                                  {fItem.cargo && <span className="font-sans text-stone-600 text-sm leading-snug">{fItem.cargo}</span>}
                                  {fItem.infoAdicional && <span className="font-sans text-stone-500 text-xs italic leading-snug">{fItem.infoAdicional}</span>}
                                </div>
                              ))}
                            </div>
                          )}

                          {/* RENDERIZADO DEL BLOQUE BOTÓN CON TAMAÑO Y FUENTE PERSONALIZADOS */}
                          {bloque.tipo === 'boton' && (() => {
                            const sz = obtenerEstilosTamanoBoton(bloque.tamanoBoton);
                            const fnt = NOMBRES_FUENTES_WORD[bloque.fuenteBoton || 'arial']?.family || 'sans-serif';
                            return (
                              <div 
                                className={`py-4 flex ${bloque.alineacion === 'centro' ? 'justify-center' : bloque.alineacion === 'derecha' ? 'justify-end' : 'justify-start'}`}
                                style={{ backgroundColor: bloque.colorFondo === 'transparent' ? 'transparent' : (bloque.colorFondo || 'transparent') }}
                              >
                                <a
                                  href={bloque.url || '#'}
                                  target={bloque.url?.startsWith('#') ? undefined : (bloque.abrirNuevaPestana !== false ? "_blank" : "_self")}
                                  rel={bloque.url?.startsWith('#') ? undefined : "noreferrer"}
                                  onClick={(e) => {
                                    if (bloque.url?.startsWith('#')) {
                                      e.preventDefault();
                                      irAPaginaInterna(bloque.url);
                                    }
                                  }}
                                  style={{
                                    backgroundColor: bloque.estilo === 'contorno' ? 'transparent' : bloque.colorBoton,
                                    color: bloque.estilo === 'contorno' ? bloque.colorBoton : bloque.colorTexto,
                                    borderColor: bloque.colorBoton,
                                    padding: sz.padding,
                                    fontSize: sz.fontSize,
                                    fontFamily: fnt,
                                    boxShadow: bloque.estilo === '3d' 
                                      ? `0 6px 0 0 ${bloque.colorBoton}aa, 0 10px 15px -3px rgba(0,0,0,0.2)` 
                                      : bloque.estilo === '4d'
                                      ? `0 0 25px 4px ${bloque.colorIluminacion || bloque.colorBoton}, inset 0 0 12px rgba(255,255,255,0.4)`
                                      : undefined
                                  }}
                                  className={`font-black uppercase tracking-wider transition-all duration-200 inline-block hover:scale-105 active:scale-95 ${
                                    bloque.estilo === 'pildora' ? 'rounded-full shadow-lg' :
                                    bloque.estilo === 'contorno' ? 'rounded-xl border-2 hover:bg-black/5' :
                                    bloque.estilo === '3d' ? 'rounded-xl -translate-y-1 active:translate-y-0.5' :
                                    bloque.estilo === '4d' ? 'rounded-full border border-white/40' :
                                    'rounded-xl shadow-lg'
                                  }`}
                                >
                                  {bloque.texto || 'Explorar'}
                                </a>
                              </div>
                            );
                          })()}

                          {bloque.tipo === 'imagen' && bloque.url && (
                            <div className={bloque.pantallaCompleta ? "w-screen relative left-1/2 -translate-x-1/2 my-12 shadow-sm" : "w-full rounded-3xl overflow-hidden shadow-sm border border-stone-200/50"} style={{ backgroundColor: bloque.colorFondo === 'transparent' ? 'transparent' : (bloque.colorFondo || 'transparent') }}>
                              <img 
                                src={bloque.url} 
                                alt="Imagen de la publicación" 
                                className={`w-full object-cover ${bloque.pantallaCompleta ? 'h-auto max-h-screen' : 'h-auto max-h-[75vh]'}`} 
                              />
                            </div>
                          )}

                          {bloque.tipo === 'tabla' && !bloque.oculto && (() => {
                            const filtroActual = (filtrosTablas[bloque.id] || '').trim();
                            const filtroNorm = quitarAcentos(filtroActual);
                            const filasCuerpoOriginales = bloque.data.slice(1);
                            const filasCuerpo = filtroNorm
                              ? filasCuerpoOriginales.filter(f => f.some(celda => quitarAcentos(String(celda)).includes(filtroNorm)))
                              : filasCuerpoOriginales;

                            return (
                              <div 
                                className="border border-stone-200/50 rounded-2xl overflow-hidden shadow-sm transition-all mx-auto" 
                                style={{ 
                                  backgroundColor: bloque.colorFondo || '#ffffff',
                                  width: `${Math.min(100, bloque.escala || 100)}%`,
                                  zoom: (bloque.escala || 100) > 100 ? (bloque.escala || 100) / 100 : 1
                                }}
                              >
                                <div className="bg-stone-50/70 border-b border-stone-200/60 px-6 py-4 flex flex-wrap justify-between items-center gap-3">
                                  <h3 className="font-sans font-bold uppercase tracking-widest text-xs text-stone-800">{bloque.nombre}</h3>
                                  
                                  <div className="flex flex-wrap items-center gap-2 font-sans">
                                    {filasCuerpoOriginales.length > 6 && (
                                      <div className="flex items-center gap-1.5 bg-white border border-stone-200 rounded-lg px-2.5 py-1">
                                        <Search size={12} className="text-stone-400" />
                                        <input
                                          type="text"
                                          value={filtrosTablas[bloque.id] || ''}
                                          onChange={(e) => setFiltrosTablas(prev => ({ ...prev, [bloque.id]: e.target.value }))}
                                          placeholder="Filtrar fila..."
                                          className="text-[11px] font-medium bg-transparent outline-none w-24 sm:w-32 text-stone-700"
                                        />
                                        {filtroActual && (
                                          <button
                                            type="button"
                                            onClick={() => setFiltrosTablas(prev => ({ ...prev, [bloque.id]: '' }))}
                                            className="text-stone-400 hover:text-stone-700"
                                          >
                                            <X size={11} />
                                          </button>
                                        )}
                                      </div>
                                    )}

                                    <button
                                      type="button"
                                      onClick={() => copiarTablaPortapapeles(bloque)}
                                      className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white hover:bg-stone-100 border border-stone-200 text-[11px] font-bold text-stone-600 transition-colors cursor-pointer"
                                      title="Copiar datos para pegar en Excel"
                                    >
                                      {tablaCopiadaId === bloque.id ? <Check size={12} className="text-emerald-600" /> : <Copy size={12} />}
                                      <span>{tablaCopiadaId === bloque.id ? 'Copiado' : 'Copiar'}</span>
                                    </button>

                                    <button
                                      type="button"
                                      onClick={() => descargarTablaCsv(bloque)}
                                      className="flex items-center gap-1 px-2.5 py-1 rounded-lg bg-white hover:bg-emerald-50 border border-stone-200 hover:border-emerald-300 text-[11px] font-bold text-stone-700 hover:text-emerald-800 transition-colors cursor-pointer"
                                      title="Descargar esta tabla en formato CSV / Excel"
                                    >
                                      <Download size={12} />
                                      <span>Excel / CSV</span>
                                    </button>
                                  </div>
                                </div>

                                <div className="overflow-x-auto">
                                  <table 
                                    className="w-full text-left font-sans border-collapse" 
                                    style={{ 
                                      backgroundColor: bloque.colorFondo || '#ffffff',
                                      fontSize: `${bloque.tamanoTextoTabla ?? 14}px`
                                    }}
                                  >
                                    <thead className="border-b border-stone-200/50 sticky top-0 z-10" style={{ backgroundColor: bloque.coloresFilas?.[0] || bloque.colorFondo || '#f5f5f4' }}>
                                      <tr>
                                        {bloque.data[0]?.map((cabecera, i) => {
                                          const padGeneral = bloque.paddingFila ?? 6;
                                          const altoCabecera = bloque.altosFilas?.[0];
                                          const fontSizeHeader = Math.max(9, (bloque.tamanoTextoTabla ?? 14) - 1);
                                          const padY = altoCabecera !== undefined && altoCabecera > 0
                                            ? Math.max(0, Math.round((altoCabecera - fontSizeHeader) / 2))
                                            : padGeneral;
                                          const anchoCol = bloque.anchosColumnas?.[i];
                                          return (
                                            <th 
                                              key={i} 
                                              className="font-bold px-4 uppercase tracking-wider align-middle"
                                              style={{ 
                                                backgroundColor: bloque.coloresFilas?.[0] || bloque.coloresColumnas?.[i] || 'transparent',
                                                fontSize: `${fontSizeHeader}px`,
                                                lineHeight: 1.1,
                                                paddingTop: `${padY}px`,
                                                paddingBottom: `${padY}px`,
                                                ...(altoCabecera ? { height: `${altoCabecera}px` } : {}),
                                                ...(anchoCol ? { width: `${anchoCol}px`, minWidth: `${anchoCol}px`, maxWidth: `${anchoCol}px` } : {})
                                              }}
                                            >
                                              {cabecera}
                                            </th>
                                          );
                                        })}
                                      </tr>
                                    </thead>
                                    <tbody className="divide-y divide-stone-100/50">
                                      {filasCuerpo.map((fila, fIdx) => {
                                        const realRowIndex = fIdx + 1;
                                        const colorFila = bloque.coloresFilas?.[realRowIndex];
                                        const padGeneral = bloque.paddingFila ?? 6;
                                        const altoFila = bloque.altosFilas?.[realRowIndex];
                                        const fontSizeBody = bloque.tamanoTextoTabla ?? 14;
                                        const padY = altoFila !== undefined && altoFila > 0
                                          ? Math.max(0, Math.round((altoFila - fontSizeBody) / 2))
                                          : padGeneral;

                                        return (
                                          <tr 
                                            key={fIdx} 
                                            className="hover:bg-black/5 transition-colors"
                                            style={{ 
                                              backgroundColor: colorFila || bloque.colorFondo || '#ffffff',
                                              ...(altoFila ? { height: `${altoFila}px` } : {})
                                            }}
                                          >
                                            {fila.map((celda, cIdx) => {
                                              const colorCol = bloque.coloresColumnas?.[cIdx];
                                              const anchoCol = bloque.anchosColumnas?.[cIdx];
                                              return (
                                                <td 
                                                  key={cIdx} 
                                                  className={`px-4 align-middle ${cIdx === 0 ? 'font-semibold' : 'opacity-85'}`}
                                                  style={{ 
                                                    backgroundColor: colorFila || colorCol || 'transparent',
                                                    fontSize: `${fontSizeBody}px`,
                                                    lineHeight: 1.1,
                                                    paddingTop: `${padY}px`,
                                                    paddingBottom: `${padY}px`,
                                                    ...(altoFila ? { height: `${altoFila}px` } : {}),
                                                    ...(anchoCol ? { width: `${anchoCol}px`, minWidth: `${anchoCol}px`, maxWidth: `${anchoCol}px` } : {})
                                                  }}
                                                >
                                                  {celda}
                                                </td>
                                              );
                                            })}
                                          </tr>
                                        );
                                      })}
                                    </tbody>
                                  </table>
                                </div>
                                {!textoVacioQuill(bloque.pieBloque) && (
                                  <div 
                                    className="px-6 py-3 bg-stone-50/40 border-t border-stone-200/50 prose prose-sm prose-stone max-w-none text-stone-600 leading-snug"
                                    dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(bloque.pieBloque || '') }}
                                  />
                                )}
                              </div>
                            );
                          })()}

                          {bloque.tipo === 'grafica' && (
                            <div 
                              className="border border-stone-200/50 rounded-2xl overflow-hidden shadow-sm p-8 transition-all mx-auto" 
                              style={{ 
                                backgroundColor: bloque.colorFondo || '#ffffff',
                                width: `${Math.min(100, bloque.escala || 100)}%`,
                                zoom: (bloque.escala || 100) > 100 ? (bloque.escala || 100) / 100 : 1
                              }}
                            >
                              <h3 className="font-sans font-bold uppercase tracking-widest text-sm text-center mb-8">{bloque.nombre}</h3>
                              {renderizarGrafica(bloque)}
                              {!textoVacioQuill(bloque.pieBloque) && (
                                <div 
                                  className="mt-4 pt-3 border-t border-stone-200/50 prose prose-sm prose-stone max-w-none text-stone-600 leading-snug"
                                  dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(bloque.pieBloque || '') }}
                                />
                              )}
                            </div>
                          )}

                        </div>
                      ))}
                    </div>

                    {pieMostrar && (
                      <div 
                        className="mt-12 pt-6 pb-24 border-t border-stone-200/60 prose prose-sm prose-stone max-w-none opacity-75"
                        dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(pieMostrar) }}
                      />
                    )}

                  </div>
                </div>
              );
            })()}
          </motion.div>
        </AnimatePresence>
      )}
      
      {/* FLECHAS LATERALES DE NAVEGACIÓN RÁPIDA */}
      {!generandoPdfPublico && (
        <AnimatePresence>
          {paginaActual > 0 && (
            <motion.button
              initial={{ opacity: 0, x: -20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -20 }}
              onClick={() => navegarASeccion(paginaActual - 1)}
              className="absolute left-6 md:left-12 top-1/2 -translate-y-1/2 w-14 h-14 md:w-16 md:h-16 bg-white/80 backdrop-blur-xl border border-stone-200 rounded-full hidden md:flex items-center justify-center text-stone-500 hover:text-blue-700 hover:border-blue-700 hover:bg-white shadow-lg shadow-stone-200/50 transition-all z-40 group cursor-pointer no-print"
            >
              <ChevronLeft size={32} strokeWidth={1.5} className="group-hover:-translate-x-1 transition-transform" />
            </motion.button>
          )}
        </AnimatePresence>
      )}

      {!generandoPdfPublico && (
        <AnimatePresence>
          {paginaActual < totalSecciones - 1 && (
            <motion.button
              initial={{ opacity: 0, x: 20 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: 20 }}
              onClick={() => navegarASeccion(paginaActual + 1)}
              className="absolute right-6 md:right-12 top-1/2 -translate-y-1/2 w-14 h-14 md:w-16 md:h-16 bg-white/80 backdrop-blur-xl border border-stone-200 rounded-full hidden md:flex items-center justify-center text-stone-500 hover:text-blue-700 hover:border-blue-700 hover:bg-white shadow-lg shadow-stone-200/50 transition-all z-40 group cursor-pointer no-print"
            >
              <ChevronRight size={32} strokeWidth={1.5} className="group-hover:translate-x-1 transition-transform" />
            </motion.button>
          )}
        </AnimatePresence>
      )}

      {/* DOCK FLOTANTE COMPACTO DE NAVEGACIÓN */}
      {!generandoPdfPublico && (
        <div className="fixed bottom-4 left-1/2 -translate-x-1/2 z-40 w-[95%] max-w-2xl pointer-events-auto no-print">
          <div className="bg-white/95 backdrop-blur-xl border border-stone-200/90 shadow-xl shadow-stone-900/10 rounded-2xl px-4 py-2.5 flex flex-col gap-2 font-sans">
            
            <div className="flex items-center justify-between gap-2">
              
              <div className="flex items-center gap-1 min-w-0">
                <button
                  type="button"
                  onClick={() => navegarASeccion(paginaActual - 1)}
                  disabled={paginaActual === 0}
                  className="p-1.5 rounded-lg hover:bg-stone-100 text-stone-600 disabled:opacity-30 transition-colors cursor-pointer shrink-0"
                  title="Página anterior"
                >
                  <ChevronLeft size={18} />
                </button>

                <div className="relative flex items-center min-w-0">
                  <ListOrdered size={14} className="text-stone-400 mr-1.5 hidden sm:inline shrink-0" />
                  <select
                    value={paginaActual}
                    onChange={(e) => navegarASeccion(Number(e.target.value))}
                    className="bg-transparent text-xs font-bold text-stone-800 outline-none cursor-pointer max-w-40 sm:max-w-60 truncate py-1 pr-2 hover:text-blue-700 transition-colors"
                    title="Saltar directamente a cualquier sección o página"
                  >
                    {secciones.map((sec, idx) => {
                      const itemIdx = indiceBusqueda[idx];
                      return (
                        <option key={idx} value={idx}>
                          {sec.tipo === 'portada'
                            ? 'Portada Principal'
                            : `Pág. ${sec.numeroReal} — ${itemIdx?.titulo || `Hoja ${sec.numeroReal}`}`}
                        </option>
                      );
                    })}
                  </select>
                </div>

                <button
                  type="button"
                  onClick={() => navegarASeccion(paginaActual + 1)}
                  disabled={paginaActual === totalSecciones - 1}
                  className="p-1.5 rounded-lg hover:bg-stone-100 text-stone-600 disabled:opacity-30 transition-colors cursor-pointer shrink-0"
                  title="Página siguiente"
                >
                  <ChevronRight size={18} />
                </button>
              </div>

              <div className="flex items-center gap-2 shrink-0">
                <form
                  onSubmit={(e) => {
                    e.preventDefault();
                    const num = Number(inputIrAPagina);
                    if (!isNaN(num) && num >= 1 && num <= paginas.length) {
                      irAPaginaInterna(String(num));
                      setInputIrAPagina('');
                    }
                  }}
                  className="flex items-center gap-1.5 bg-stone-100/90 border border-stone-200/90 rounded-xl px-3 py-1"
                >
                  <span className="text-[10px] font-bold uppercase tracking-wider text-stone-400 hidden sm:inline">Ir a pág:</span>
                  <input
                    type="text"
                    inputMode="numeric"
                    value={inputIrAPagina}
                    onChange={(e) => setInputIrAPagina(e.target.value.replace(/[^0-9]/g, ''))}
                    placeholder={seccionActiva.tipo === 'portada' ? '0' : `${seccionActiva.numeroReal}`}
                    className="w-9 text-center text-xs font-black text-stone-900 bg-white border border-stone-200/80 rounded-md py-0.5 outline-none focus:border-blue-500 transition-colors"
                    title="Escribe un número de página y presiona Enter"
                  />
                  <span className="text-xs font-black text-stone-400 whitespace-nowrap">
                    / {paginas.length}
                  </span>
                  {inputIrAPagina && (
                    <button
                      type="submit"
                      className="text-blue-600 hover:text-blue-800 p-0.5 cursor-pointer"
                      title="Saltar a la página"
                    >
                      <ArrowRight size={13} />
                    </button>
                  )}
                </form>

                <button
                  type="button"
                  onClick={() => setBuscadorAbierto(true)}
                  className="p-2 rounded-xl bg-stone-900 hover:bg-blue-700 text-white transition-colors cursor-pointer shadow-xs"
                  title="Buscar en todo el libro (Ctrl + K)"
                >
                  <Search size={14} />
                </button>
              </div>
            </div>

            <div className="flex items-center gap-3 px-1">
              <span className="text-[10px] font-bold uppercase tracking-widest text-stone-400 shrink-0 min-w-12">
                {seccionActiva.tipo === 'portada' ? 'Portada' : `Pág ${seccionActiva.numeroReal}`}
              </span>
              <input
                type="range"
                min={0}
                max={Math.max(0, totalSecciones - 1)}
                value={paginaActual}
                onChange={(e) => navegarASeccion(Number(e.target.value))}
                style={{ accentColor: edicionActiva?.color_principal || '#1d4ed8' }}
                className="w-full h-1.5 bg-stone-200 rounded-lg cursor-pointer"
                title="Arrastra para desplazarte rápidamente a cualquier punto del libro"
              />
              <span className="text-[10px] font-bold uppercase tracking-widest text-stone-400 shrink-0">
                {paginas.length} págs
              </span>
            </div>

          </div>
        </div>
      )}

      {/* MODAL DEL BUSCADOR DINÁMICO INTELIGENTE */}
      <AnimatePresence>
        {buscadorAbierto && !generandoPdfPublico && (
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.15 }}
            className="fixed inset-0 z-50 flex items-start justify-center pt-10 sm:pt-16 px-4 font-sans no-print"
          >
            <div
              className="absolute inset-0 bg-stone-900/45 backdrop-blur-xs"
              onClick={() => setBuscadorAbierto(false)}
            />

            <motion.div
              initial={{ opacity: 0, scale: 0.96, y: -10 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.96, y: -10 }}
              transition={{ duration: 0.15 }}
              className="relative z-10 w-full max-w-2xl bg-[#fdfbf7] border border-stone-200/90 rounded-3xl shadow-2xl overflow-hidden"
            >
              <div className="flex items-center gap-3 px-6 py-4 bg-white border-b border-stone-200/80">
                <Search size={20} style={{ color: edicionActiva.color_principal || '#1d4ed8' }} className="shrink-0" />
                <input
                  type="text"
                  autoFocus
                  value={terminoBusqueda}
                  onChange={(e) => setTerminoBusqueda(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === 'ArrowDown') {
                      e.preventDefault();
                      setIndiceResultadoSeleccionado(prev => Math.min(resultadosBusqueda.length - 1, prev + 1));
                    } else if (e.key === 'ArrowUp') {
                      e.preventDefault();
                      setIndiceResultadoSeleccionado(prev => Math.max(0, prev - 1));
                    } else if (e.key === 'Enter' && resultadosBusqueda[indiceResultadoSeleccionado]) {
                      e.preventDefault();
                      navegarASeccion(resultadosBusqueda[indiceResultadoSeleccionado].idxSeccion);
                      setBuscadorAbierto(false);
                    }
                  }}
                  placeholder="Busca una ingeniería (ej. Sistemas, Civil), ciudad o escribe «pagina 5»..."
                  className="w-full bg-transparent text-stone-900 text-sm sm:text-base font-medium outline-none placeholder:text-stone-400"
                />
                {terminoBusqueda && (
                  <button
                    type="button"
                    onClick={() => setTerminoBusqueda('')}
                    className="text-xs font-bold text-stone-400 hover:text-stone-700 px-2.5 py-1 rounded-lg hover:bg-stone-100 transition-colors cursor-pointer"
                  >
                    Limpiar
                  </button>
                )}
                <button
                  type="button"
                  onClick={() => setBuscadorAbierto(false)}
                  className="p-1.5 rounded-xl bg-stone-100 hover:bg-stone-200 text-stone-600 transition-colors cursor-pointer"
                >
                  <X size={18} />
                </button>
              </div>

              <div className="px-6 py-3 bg-stone-100/70 border-b border-stone-200/70 flex flex-wrap items-center justify-between gap-2">
                <div className="flex items-center gap-1.5">
                  <button
                    type="button"
                    onClick={() => setFiltroCategoriaBusqueda('todo')}
                    className={`px-3 py-1 rounded-full text-xs font-bold transition-all cursor-pointer ${
                      filtroCategoriaBusqueda === 'todo'
                        ? 'bg-stone-900 text-white shadow-xs'
                        : 'bg-white text-stone-600 hover:bg-stone-200/70 border border-stone-200/80'
                    }`}
                  >
                    Todo el libro
                  </button>
                  <button
                    type="button"
                    onClick={() => setFiltroCategoriaBusqueda('capitulos')}
                    className={`px-3 py-1 rounded-full text-xs font-bold transition-all flex items-center gap-1 cursor-pointer ${
                      filtroCategoriaBusqueda === 'capitulos'
                        ? 'bg-stone-900 text-white shadow-xs'
                        : 'bg-white text-stone-600 hover:bg-stone-200/70 border border-stone-200/80'
                    }`}
                  >
                    <FileText size={12} /> Capítulos
                  </button>
                  <button
                    type="button"
                    onClick={() => setFiltroCategoriaBusqueda('tablas')}
                    className={`px-3 py-1 rounded-full text-xs font-bold transition-all flex items-center gap-1 cursor-pointer ${
                      filtroCategoriaBusqueda === 'tablas'
                        ? 'bg-stone-900 text-white shadow-xs'
                        : 'bg-white text-stone-600 hover:bg-stone-200/70 border border-stone-200/80'
                    }`}
                  >
                    <TableIcon size={12} /> Tablas y Gráficas
                  </button>
                </div>

                <div className="hidden sm:flex items-center gap-1.5 text-[11px]">
                  <span className="text-stone-400 font-semibold">Rápido:</span>
                  {['Tabla de contenido', 'Sistemas', 'Industrial', 'Civil', 'Posgrado'].map(sug => (
                    <button
                      key={sug}
                      type="button"
                      onClick={() => setTerminoBusqueda(sug)}
                      className="px-2 py-0.5 bg-white hover:bg-blue-50 text-stone-600 hover:text-blue-700 border border-stone-200 rounded-md font-semibold transition-colors cursor-pointer"
                    >
                      {sug}
                    </button>
                  ))}
                </div>
              </div>

              <div className="max-h-[58vh] overflow-y-auto divide-y divide-stone-200/40 p-3 space-y-1.5 custom-scrollbar">
                {resultadosBusqueda.length === 0 ? (
                  <div className="py-14 text-center text-stone-400">
                    <p className="text-base font-bold text-stone-600 mb-1">No se encontraron coincidencias para &ldquo;{terminoBusqueda}&rdquo;</p>
                    <p className="text-xs">Prueba escribiendo sin ortografía estricta, el nombre de una ingeniería o el número de página (ej. &ldquo;pagina 15&rdquo;).</p>
                  </div>
                ) : (
                  resultadosBusqueda.map((item: any, rIdx: number) => {
                    const esActiva = item.idxSeccion === paginaActual;
                    const esSeleccionadaTeclado = rIdx === indiceResultadoSeleccionado;
                    return (
                      <button
                        key={item.idxSeccion}
                        type="button"
                        onMouseEnter={() => setIndiceResultadoSeleccionado(rIdx)}
                        onClick={() => {
                          navegarASeccion(item.idxSeccion);
                          setBuscadorAbierto(false);
                        }}
                        className={`w-full text-left p-4 rounded-2xl transition-all flex items-start justify-between gap-4 cursor-pointer ${
                          esSeleccionadaTeclado
                            ? 'bg-white ring-2 ring-blue-500/40 shadow-md border border-blue-200'
                            : esActiva
                            ? 'bg-blue-50/70 border border-blue-200'
                            : 'bg-white/75 hover:bg-white border border-transparent hover:border-stone-200/80'
                        }`}
                      >
                        <div className="min-w-0 flex-1">
                          <div className="flex flex-wrap items-center gap-2 mb-1.5">
                            <span
                              className="px-2.5 py-0.5 text-[10px] font-black uppercase tracking-wider rounded-full text-white shrink-0"
                              style={{ backgroundColor: item.esPortada ? '#57534e' : (edicionActiva.color_principal || '#1d4ed8') }}
                            >
                              {item.esPortada ? 'Portada' : `Página ${item.numeroReal}`}
                            </span>

                            {item.esCoincidenciaExactaPagina && (
                              <span className="px-2 py-0.5 text-[10px] font-black uppercase tracking-wider rounded-full bg-amber-100 text-amber-800 border border-amber-300 flex items-center gap-1">
                                <Sparkles size={10} /> Salto directo a página
                              </span>
                            )}

                            {item.tieneTabla && (
                              <span className="px-2 py-0.5 text-[10px] font-bold rounded-md bg-stone-100 text-stone-600 flex items-center gap-1">
                                <TableIcon size={10} /> Tabla
                              </span>
                            )}

                            {item.tieneGrafica && (
                              <span className="px-2 py-0.5 text-[10px] font-bold rounded-md bg-stone-100 text-stone-600 flex items-center gap-1">
                                <BarChart3 size={10} /> Gráfica
                              </span>
                            )}
                          </div>

                          <h4 className="font-serif font-bold text-stone-900 text-base truncate">
                            {resaltarCoincidencia(item.titulo, item.terminoResaltar)}
                          </h4>

                          {item.extracto && (
                            <p className="text-xs text-stone-500 line-clamp-2 leading-relaxed mt-1">
                              {resaltarCoincidencia(item.extracto, item.terminoResaltar)}
                            </p>
                          )}
                        </div>
                        <div className="shrink-0 self-center flex items-center gap-1 text-xs font-bold text-stone-400">
                          <span className="hidden sm:inline text-[10px] uppercase tracking-wider">Ir</span>
                          <ArrowRight size={16} />
                        </div>
                      </button>
                    );
                  })
                )}
              </div>

              <div className="px-6 py-3 bg-stone-100/90 border-t border-stone-200/70 flex items-center justify-between text-[11px] text-stone-500">
                <span>
                  Mostrando <strong>{resultadosBusqueda.length}</strong> {resultadosBusqueda.length === 1 ? 'resultado relevante' : 'resultados relevantes'}
                </span>
                <div className="hidden sm:flex items-center gap-3">
                  <span><strong>↑↓</strong> Navegar</span>
                  <span><strong>Enter</strong> Ir a página</span>
                  <span><strong>ESC</strong> Cerrar</span>
                </div>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

    </div>
  );
}

export default function LibroPublico() {
  return (
    <Suspense fallback={
      <div className="flex h-screen items-center justify-center bg-[#fdfbf7]">
        <div className="w-16 h-16 border-4 border-stone-200 border-t-blue-700 rounded-full animate-spin"></div>
      </div>
    }>
      <LibroPublicoContent />
    </Suspense>
  );
}