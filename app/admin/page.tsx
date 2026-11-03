'use client';

import React, { useEffect, useState, useMemo, useRef } from 'react';
import dynamic from 'next/dynamic';
import { supabase } from '../../lib/supabase';
import { 
  Settings, Save, CheckCircle2, AlertCircle, Copy, Eye, EyeOff,
  Edit3, Power, Database, Layout, Trash2, Book, Type, RefreshCcw, 
  Plus, Minus, ChevronLeft, ChevronRight, GripVertical, Table as TableIcon, 
  X, BarChart3, Map as MapIcon, Image as ImageIcon, Palette,
  ArrowUp, ArrowDown, MoveLeft, MoveRight, Upload, Lock, ImagePlus, MousePointerClick,
  PanelLeftClose, PanelLeftOpen, Link2, PenTool, Columns, Search, Download, Printer, FileJson,
  AlignLeft, AlignCenter, AlignRight, Square
} from 'lucide-react';

import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer, 
  PieChart, Pie, Cell, LineChart, Line, AreaChart, Area, RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar,
  ComposedChart, LabelList, RadialBarChart, RadialBar, FunnelChart, Funnel, Treemap, Sector, ScatterChart, Scatter
} from 'recharts';

import 'react-quill-new/dist/quill.snow.css';

// LISTA COMPLETA DE FUENTES ESTILO MICROSOFT WORD
const FUENTES_WORD = [
  'arial',
  'arial-black',
  'bahnschrift',
  'calibri',
  'cambria',
  'candara',
  'century-gothic',
  'comic-sans-ms',
  'consolas',
  'constantia',
  'corbel',
  'courier-new',
  'franklin-gothic-medium',
  'garamond',
  'georgia',
  'impact',
  'lucida-console',
  'lucida-sans-unicode',
  'palatino-linotype',
  'segoe-ui',
  'tahoma',
  'times-new-roman',
  'trebuchet-ms',
  'verdana'
];

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

const ESTILOS_FUENTES_QUILL = FUENTES_WORD.map(f => `
  .ql-Font-${f}, .ql-font-${f} { font-family: ${NOMBRES_FUENTES_WORD[f].family}; }
  .ql-snow .ql-picker.ql-font .ql-picker-label[data-value="${f}"]::before,
  .ql-snow .ql-picker.ql-font .ql-picker-item[data-value="${f}"]::before {
    content: "${NOMBRES_FUENTES_WORD[f].label}";
    font-family: ${NOMBRES_FUENTES_WORD[f].family};
  }
`).join('\n') + `
  .ql-snow .ql-picker.ql-font {
    width: 155px;
  }
  .ql-snow .ql-picker.ql-font .ql-picker-options {
    max-height: 260px;
    overflow-y: auto;
    z-index: 9999;
    border-radius: 12px;
    box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.2);
  }
  .ql-toolbar.ql-snow {
    position: relative;
    z-index: 30;
  }
  .ql-editor img {
    max-width: 100%;
    height: auto;
    cursor: pointer;
    transition: box-shadow 0.2s ease;
  }
  .ql-editor img:hover {
    box-shadow: 0 0 0 3px #3b82f6;
  }
  .ql-editor .ql-align-center {
    text-align: center;
  }
  .ql-editor .ql-align-center img {
    margin-left: auto;
    margin-right: auto;
    display: block;
  }
  .ql-editor .ql-align-right {
    text-align: right;
  }
  .ql-editor .ql-align-right img {
    margin-left: auto;
    margin-right: 0;
    display: block;
  }
  .ql-snow .ql-editor a {
    color: inherit;
  }
  .prose p {
    margin-top: 0 !important;
    margin-bottom: 0.25rem !important;
    min-height: 1.5em;
    line-height: 1.5 !important;
  }
  /* Clases dedicadas para imágenes personalizadas dentro del texto */
  img[data-borde="cuadrado"], .ql-editor img[data-borde="cuadrado"] {
    border-radius: 0px !important;
  }
  img[data-borde="redondeado"], .ql-editor img[data-borde="redondeado"] {
    border-radius: 14px !important;
  }
  img[data-borde="circular"], .ql-editor img[data-borde="circular"] {
    border-radius: 9999px !important;
  }
  img[data-alineacion="izquierda"], .ql-editor img[data-alineacion="izquierda"] {
    display: block !important;
    margin-left: 0 !important;
    margin-right: auto !important;
  }
  img[data-alineacion="centro"], .ql-editor img[data-alineacion="centro"] {
    display: block !important;
    margin-left: auto !important;
    margin-right: auto !important;
  }
  img[data-alineacion="derecha"], .ql-editor img[data-alineacion="derecha"] {
    display: block !important;
    margin-left: auto !important;
    margin-right: 0 !important;
  }
  @media print {
    @page {
      size: A4 landscape;
      margin: 10mm;
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
      page-break-after: always !important;
      break-after: page !important;
      min-height: 185mm !important;
      width: 100% !important;
      padding: 12mm 14mm !important;
      box-sizing: border-box !important;
      position: relative !important;
      overflow: visible !important;
    }
    .print-page-sheet:last-child {
      page-break-after: auto !important;
      break-after: auto !important;
    }
    .print-avoid-break {
      page-break-inside: avoid !important;
      break-inside: avoid !important;
    }
  }
`;

const ReactQuill = dynamic(
  async () => {
    const mod = await import('react-quill-new');
    const Quill = mod.default.Quill || (mod as any).Quill;
    if (Quill) {
      const Font = Quill.import('formats/font') as any;
      if (Font) {
        Font.whitelist = FUENTES_WORD;
        Quill.register(Font, true);
      }
      // Permitir atributos personalizados en imágenes dentro de Quill (width, style, data-borde, data-alineacion)
      const BaseImage = Quill.import('formats/image') as any;
      if (BaseImage) {
        const ATTRIBUTES = ['alt', 'height', 'width', 'style', 'data-borde', 'data-alineacion'];
        class CustomImage extends BaseImage {
          static formats(domNode: HTMLElement) {
            return ATTRIBUTES.reduce((formats: Record<string, string | null>, attribute) => {
              if (domNode.hasAttribute(attribute)) {
                formats[attribute] = domNode.getAttribute(attribute);
              }
              return formats;
            }, {});
          }
          format(name: string, value: any) {
            if (ATTRIBUTES.indexOf(name) > -1) {
              if (value) {
                this.domNode.setAttribute(name, value);
              } else {
                this.domNode.removeAttribute(name);
              }
            } else {
              super.format(name, value);
            }
          }
        }
        Quill.register(CustomImage, true);
      }
    }
    return mod;
  },
  { ssr: false }
);

import { ComposableMap, Geographies, Geography, ZoomableGroup } from 'react-simple-maps';
import { scaleLinear } from 'd3-scale';

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
  mostrarValores?: boolean; 
  pieBloque?: string;
};

type BloqueImagen = BloqueBase & { tipo: 'imagen'; url: string; alt: string; pantallaCompleta: boolean; };

type EstiloBoton = 'solido' | 'contorno' | 'pildora' | '3d' | '4d';
type BloqueBoton = BloqueBase & {
  tipo: 'boton';
  texto: string;
  url: string;
  estilo: EstiloBoton;
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
};

type ConfiguracionEdicion = {
  subtitulo?: string;
  portadaContenido?: string;
  colorFondoPortada?: string;
  imagenFondoPortada?: string;
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

export default function AdminPage() {
  const [sesion, setSesion] = useState<any>(null);
  const [revisandoSesion, setRevisandoSesion] = useState(true);
  
  const [emailLogin, setEmailLogin] = useState('');
  const [passwordLogin, setPasswordLogin] = useState('');
  const [cargandoLogin, setCargandoLogin] = useState(false);
  const [errorLogin, setErrorLogin] = useState('');

  const [ediciones, setEdiciones] = useState<Edicion[]>([]);
  const [edicionActiva, setEdicionActiva] = useState<Edicion | null>(null);
  const [pestañaActiva, setPestañaActiva] = useState<'diseno' | 'contenido'>('contenido');

  const [menuLateralAbierto, setMenuLateralAbierto] = useState(true);

  const [paginas, setPaginas] = useState<Pagina[]>([]);
  const [configuracion, setConfiguracion] = useState<ConfiguracionEdicion>({ 
    modoMarcaAgua: 'todas', 
    paginasMarcaAgua: [],
    colorFondoPortada: '#fdfbf7',
    logoSize: 56,
    logoBrightness: 100,
    marcaAguaSize: 40,
    marcaAguaOpacity: 5,
    cabeceraGlobalHtml: '',
    piePaginaGlobalHtml: '',
    mostrarCabeceraGlobal: false,
    mostrarPiePaginaGlobal: false
  });

  const [indicePaginaActual, setIndicePaginaActual] = useState<number>(0);

  const [cargando, setCargando] = useState(false); 
  const [guardando, setGuardando] = useState(false);
  const [estadoGuardadoVivo, setEstadoGuardadoVivo] = useState<'guardado' | 'escribiendo' | 'guardando'>('guardado');
  const [eliminando, setEliminando] = useState(false);
  const [mensaje, setMensaje] = useState<{ texto: string; tipo: 'exito' | 'error' } | null>(null);

  const [subiendoImagen, setSubiendoImagen] = useState<Record<string, boolean>>({});
  
  const [tooltipMapa, setTooltipMapa] = useState<{ nombre: string; valor: string; x: number; y: number } | null>(null);

  const [selectorColorActivo, setSelectorColorActivo] = useState<{
    bloqueId: string;
    indice: number;
    etiqueta: string;
  } | null>(null);

  // Estado para capturar tanto el texto sombreado como la posición exacta del cursor dentro de un bloque de texto
  const [seleccionTextoActiva, setSeleccionTextoActiva] = useState<{
    bloqueId: string;
    columna: 1 | 2 | 3;
    texto: string;
  } | null>(null);
  const rangoSeleccionRef = useRef<Range | null>(null);
  const ultimoRangoCursorRef = useRef<{
    bloqueId: string;
    columna: 1 | 2 | 3;
    range: Range;
    contenedor: HTMLElement;
  } | null>(null);
  const editorContenedorRef = useRef<HTMLElement | null>(null);

  const [modalNuevaTablaAbierto, setModalNuevaTablaAbierto] = useState(false);

  const [modalImportarJsonAbierto, setModalImportarJsonAbierto] = useState(false);
  const [jsonLoteInput, setJsonLoteInput] = useState('');
  const [modoInsercionJson, setModoInsercionJson] = useState<'final' | 'despues_actual' | 'reemplazar'>('final');

  const [generandoPdfAdmin, setGenerandoPdfAdmin] = useState(false);

  const [modalConfirmacion, setModalConfirmacion] = useState<{
    abierto: boolean; titulo: string; mensaje: string; textoBotonConfirmar: string; esPeligro: boolean; accionConfirmar: (() => void) | null;
  }>({
    abierto: false, titulo: '', mensaje: '', textoBotonConfirmar: '', esPeligro: false, accionConfirmar: null
  });

  const timerRef = useRef<NodeJS.Timeout | null>(null);

  const COLORES_GRAFICAS_DEFAULT = [
    '#93c5fd', '#86efac', '#fda4af', '#d8b4fe', '#fdba74',
    '#5eead4', '#f9a8d4', '#fde047', '#a5b4fc', '#cbd5e1',
    '#bef264', '#67e8f9', '#fca5a5', '#c4b5fd', '#fed7aa'
  ];

  const quillModules = useMemo(() => ({
    toolbar: [
      [{ 'font': FUENTES_WORD }, { 'size': ['small', false, 'large', 'huge'] }],
      [{ 'header': [1, 2, 3, 4, 5, 6, false] }],
      ['bold', 'italic', 'underline', 'strike'],
      [{ 'color': [] }, { 'background': [] }],
      [{ 'script': 'sub'}, { 'script': 'super' }],
      [{ 'list': 'ordered'}, { 'list': 'bullet' }, { 'indent': '-1'}, { 'indent': '+1' }],
      [{ 'align': [] }],
      ['link', 'image', 'clean']
    ]
  }), []);

  const generarId = () => Math.random().toString(36).substr(2, 9);

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSesion(session);
      setRevisandoSesion(false);
      if (session) cargarEdiciones();
    });

    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, session) => {
      setSesion(session);
      if (session) cargarEdiciones();
    });

    return () => subscription.unsubscribe();
  }, []);

  const manejarLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setCargandoLogin(true);
    setErrorLogin('');
    
    const { error } = await supabase.auth.signInWithPassword({
      email: emailLogin,
      password: passwordLogin,
    });

    if (error) setErrorLogin(error.message === 'Invalid login credentials' ? 'Credenciales incorrectas' : error.message);
    setCargandoLogin(false);
  };

  const cerrarSesion = async () => {
    await supabase.auth.signOut();
    setEdicionActiva(null);
    setEdiciones([]);
  };

  const cambiarPagina = (nuevoIndice: number, idEdicion = edicionActiva?.id) => {
    setIndicePaginaActual(nuevoIndice);
    if (idEdicion) {
      sessionStorage.setItem(`admin_pagina_${idEdicion}`, nuevoIndice.toString());
    }
  };

  const cargarPaginasDesdeDB = (edicion: Edicion) => {
    let formatoBloques: Pagina[] = [];
    let configCargada: ConfiguracionEdicion = { 
      modoMarcaAgua: 'todas', 
      paginasMarcaAgua: [],
      colorFondoPortada: '#fdfbf7',
      logoSize: 56,
      logoBrightness: 100,
      marcaAguaSize: 40,
      marcaAguaOpacity: 5,
      cabeceraGlobalHtml: '',
      piePaginaGlobalHtml: '',
      mostrarCabeceraGlobal: false,
      mostrarPiePaginaGlobal: false
    };

    try {
      const parsed = JSON.parse(edicion.contenido_presentacion);
      if (Array.isArray(parsed)) {
        formatoBloques = parsed;
      } else if (parsed && parsed.paginas) {
        formatoBloques = parsed.paginas;
        configCargada = { ...configCargada, ...parsed.configuracion };
      } else {
        const textoCrudo = edicion.contenido_presentacion || "<h1>Página Inicial</h1>";
        const paginasDivididas = textoCrudo.split(/(?=<h1)/i).filter((p: string) => p.trim() !== '');
        formatoBloques = paginasDivididas.map((contenido: string, idx: number) => ({
            id: generarId(), 
            titulo: `Página ${idx + 1}`, 
            subtitulo: '',
            mostrarNumero: true,
            posicion: 'despues_portada',
            bloques: [{ id: generarId(), tipo: 'texto' as const, contenido }]
        }));
      }
    } catch (e) {
      formatoBloques = [{ 
        id: generarId(), 
        titulo: 'Página 1', 
        subtitulo: '',
        mostrarNumero: true,
        posicion: 'despues_portada',
        bloques: [{ id: generarId(), tipo: 'texto', contenido: edicion.contenido_presentacion || '' }] 
      }];
    }
    
    setPaginas(formatoBloques);
    setConfiguracion(configCargada);
    
    const paginaGuardada = sessionStorage.getItem(`admin_pagina_${edicion.id}`);
    if (paginaGuardada && Number(paginaGuardada) < formatoBloques.length) {
      setIndicePaginaActual(Number(paginaGuardada));
    } else {
      cambiarPagina(0, edicion.id);
    }
  };

  const cargarEdiciones = async () => {
    setCargando(true);
    try {
      const { data, error } = await supabase.from('ediciones').select('*').order('anio', { ascending: false });
      if (error) throw error;
      if (data && data.length > 0) {
        setEdiciones(data as Edicion[]);
        const edicionGuardadaId = sessionStorage.getItem('admin_edicion_activa');
        let seleccionada = data[0];
        if (edicionGuardadaId) {
          const encontrada = data.find(e => e.id.toString() === edicionGuardadaId);
          if (encontrada) seleccionada = encontrada;
        }
        setEdicionActiva(seleccionada as Edicion);
        cargarPaginasDesdeDB(seleccionada as Edicion);
      } else {
        setEdiciones([]);
        setEdicionActiva(null);
      }
    } catch (err: any) {
      mostrarMensaje(`Error: ${err.message}`, 'error');
    } finally {
      setCargando(false);
    }
  };

  const cambiarEdicionActiva = (ed: Edicion) => {
    setEdicionActiva(ed);
    sessionStorage.setItem('admin_edicion_activa', ed.id.toString());
    cargarPaginasDesdeDB(ed);
  };

  const mostrarMensaje = (texto: string, tipo: 'exito' | 'error') => {
    setMensaje({ texto, tipo });
    setTimeout(() => setMensaje(null), 4000);
  };

  const autoGuardar = (nuevasPaginas: Pagina[], nuevaConfig: ConfiguracionEdicion = configuracion) => {
    if (!edicionActiva) return;
    setEstadoGuardadoVivo('escribiendo');
    setPaginas(nuevasPaginas);
    setConfiguracion(nuevaConfig);

    if (timerRef.current) clearTimeout(timerRef.current);
    timerRef.current = setTimeout(async () => {
      setEstadoGuardadoVivo('guardando');
      try {
        const jsonToSave = JSON.stringify({ paginas: nuevasPaginas, configuracion: nuevaConfig });
        const { error } = await supabase.from('ediciones').update({ contenido_presentacion: jsonToSave }).eq('id', edicionActiva.id);
        if (error) throw error;
        setEdicionActiva({ ...edicionActiva, contenido_presentacion: jsonToSave });
        setEstadoGuardadoVivo('guardado');
      } catch (err: any) {
        setEstadoGuardadoVivo('escribiendo');
        mostrarMensaje(`Aviso al guardar en nube: ${err?.message || 'Revisa permisos RLS'}`, 'error');
      }
    }, 1500);
  };

  const descargarRespaldoJson = () => {
    if (!edicionActiva) return;
    try {
      const payload = {
        edicion: {
          titulo: edicionActiva.titulo,
          anio: edicionActiva.anio,
          color_principal: edicionActiva.color_principal,
          estado: edicionActiva.estado
        },
        configuracion,
        paginas
      };
      const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `ACOFI_Libro_Edicion_${edicionActiva.anio}_Backup.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
      mostrarMensaje(`Respaldo JSON (${paginas.length} páginas) descargado con éxito.`, 'exito');
    } catch (err: any) {
      mostrarMensaje(`Error al exportar JSON: ${err.message}`, 'error');
    }
  };

  const descargarLibroPdfAdmin = () => {
    setGenerandoPdfAdmin(true);
    setTimeout(() => {
      window.print();
      setTimeout(() => {
        setGenerandoPdfAdmin(false);
      }, 800);
    }, 1400);
  };
  // FUNCIÓN DEL INYECTOR DE LOTES JSON INTELIGENTE
  const procesarInyeccionLoteJson = () => {
    if (!jsonLoteInput.trim()) {
      mostrarMensaje('Pega primero el bloque JSON de las páginas.', 'error');
      return;
    }

    try {
      const textoLimpio = jsonLoteInput
        .trim()
        .replace(/^```(?:json)?\s*/i, '')
        .replace(/\s*```$/, '');

      const parsed = JSON.parse(textoLimpio);

      let listaPaginasCrudas: any[] = [];
      if (Array.isArray(parsed)) {
        listaPaginasCrudas = parsed;
      } else if (parsed && Array.isArray(parsed.paginas)) {
        listaPaginasCrudas = parsed.paginas;
      } else if (parsed && typeof parsed === 'object' && Array.isArray(parsed.bloques)) {
        listaPaginasCrudas = [parsed];
      } else {
        throw new Error('El JSON debe ser un arreglo de páginas [ ... ] o un objeto con la propiedad "paginas".');
      }

      if (listaPaginasCrudas.length === 0) {
        throw new Error('El lote JSON no contiene páginas válidas.');
      }

      const mapaIdsTablas: Record<string, string> = {};
      const mapaIdsPaginas: Record<string, string> = {};

      const paginasNormalizadas: Pagina[] = listaPaginasCrudas.map((pag: any, idx: number) => {
        const nuevaPagId = generarId();
        if (pag.id) {
          mapaIdsPaginas[String(pag.id)] = nuevaPagId;
        }
        const bloquesCrudos = Array.isArray(pag.bloques) ? pag.bloques : [];
        const bloquesConNuevosIds = bloquesCrudos.map((b: any) => {
          const nuevoBloqueId = generarId();
          if (b.tipo === 'tabla' && b.id) {
            mapaIdsTablas[String(b.id)] = nuevoBloqueId;
          }
          return { ...b, _oldId: b.id, id: nuevoBloqueId };
        });

        return {
          id: nuevaPagId,
          titulo: typeof pag.titulo === 'string' ? pag.titulo : `Página Importada ${idx + 1}`,
          subtitulo: typeof pag.subtitulo === 'string' ? pag.subtitulo : '',
          mostrarNumero: pag.mostrarNumero !== false,
          posicion: pag.posicion === 'antes_portada' ? 'antes_portada' : 'despues_portada',
          cabeceraHtml: pag.cabeceraHtml || '',
          piePaginaHtml: pag.piePaginaHtml || '',
          colorFondo: pag.colorFondo || '#fdfbf7',
          bloques: bloquesConNuevosIds
        };
      });

      const reemplazarLinksInternos = (htmlStr: string) => {
        if (!htmlStr) return '';
        let res = String(htmlStr);
        Object.keys(mapaIdsPaginas).forEach(oldPagId => {
          const newPagId = mapaIdsPaginas[oldPagId];
          res = res.split(`#pagina:${oldPagId}`).join(`#pagina:${newPagId}`);
          res = res.split(`#${oldPagId}`).join(`#pagina:${newPagId}`);
        });
        return res;
      };

      const paginasListas: Pagina[] = paginasNormalizadas.map((pag) => {
        const bloquesSanitizados: Bloque[] = pag.bloques.map((b: any) => {
          if (b.tipo === 'texto') {
            const cols = b.columnas === 2 || b.columnas === 3 ? b.columnas : 1;
            return {
              id: b.id,
              tipo: 'texto',
              contenido: reemplazarLinksInternos(String(b.contenido || '')),
              columnas: cols,
              contenidoCol2: reemplazarLinksInternos(String(b.contenidoCol2 || '')),
              contenidoCol3: reemplazarLinksInternos(String(b.contenidoCol3 || '')),
              colorFondo: b.colorFondo || 'transparent'
            };
          }
          if (b.tipo === 'firma') {
            const listaFirmas: ItemFirma[] = Array.isArray(b.firmas) && b.firmas.length > 0
              ? b.firmas.map((f: any) => ({
                  id: generarId(),
                  url: String(f.url || ''),
                  nombre: String(f.nombre || ''),
                  cargo: String(f.cargo || ''),
                  infoAdicional: String(f.infoAdicional || ''),
                  tamanoFirma: Number(f.tamanoFirma) || 72,
                  espacioSuperior: Number(f.espacioSuperior) || 0
                }))
              : [{
                  id: generarId(),
                  url: String(b.url || ''),
                  nombre: String(b.nombre || 'Nombre y Apellido'),
                  cargo: String(b.cargo || 'Cargo Institucional'),
                  infoAdicional: String(b.infoAdicional || ''),
                  tamanoFirma: Number(b.tamanoFirma) || 72,
                  espacioSuperior: Number(b.espacioSuperior) || 0
                }];
            return {
              id: b.id,
              tipo: 'firma',
              alineacion: b.alineacion || 'espaciado',
              espacioGeneralSuperior: Number(b.espacioGeneralSuperior) || 0,
              espacioGeneralInferior: Number(b.espacioGeneralInferior) || 0,
              firmas: listaFirmas,
              colorFondo: b.colorFondo || 'transparent'
            };
          }
          if (b.tipo === 'tabla') {
            const matrizValida = Array.isArray(b.data) && b.data.length > 0
              ? b.data.map((f: any) => (Array.isArray(f) ? f.map((c: any) => String(c ?? '')) : ['']))
              : [['Columna 1', 'Columna 2'], ['Dato 1', '0']];
            return {
              id: b.id,
              tipo: 'tabla',
              nombre: String(b.nombre || 'Tabla de Datos'),
              data: matrizValida,
              oculto: !!b.oculto,
              colorFondo: b.colorFondo || '#ffffff',
              escala: Number(b.escala) || 100,
              paddingFila: b.paddingFila !== undefined ? Number(b.paddingFila) : 6,
              tamanoTextoTabla: Number(b.tamanoTextoTabla) || 14,
              coloresFilas: b.coloresFilas || {},
              coloresColumnas: b.coloresColumnas || {},
              anchosColumnas: b.anchosColumnas || {},
              altosFilas: b.altosFilas || {},
              pieBloque: b.pieBloque || ''
            };
          }
          if (b.tipo === 'grafica') {
            const refOriginal = String(b.tablaRef || '');
            const refActualizada = mapaIdsTablas[refOriginal] || refOriginal;
            return {
              id: b.id,
              tipo: 'grafica',
              nombre: String(b.nombre || 'Gráfica Estadística'),
              tablaRef: refActualizada,
              tipoGrafica: (b.tipoGrafica as TipoGrafica) || 'barras',
              ejeX: Number(b.ejeX) || 0,
              ejesY: Array.isArray(b.ejesY) ? b.ejesY.map(Number) : [1],
              coloresCustom: Array.isArray(b.coloresCustom) ? b.coloresCustom : [],
              mostrarValores: b.mostrarValores !== false,
              colorFondo: b.colorFondo || '#ffffff',
              escala: Number(b.escala) || 100,
              pieBloque: b.pieBloque || ''
            };
          }
          if (b.tipo === 'imagen') {
            return {
              id: b.id,
              tipo: 'imagen',
              url: String(b.url || ''),
              alt: String(b.alt || ''),
              pantallaCompleta: !!b.pantallaCompleta,
              colorFondo: b.colorFondo || 'transparent'
            };
          }
          if (b.tipo === 'boton') {
            return {
              id: b.id,
              tipo: 'boton',
              texto: String(b.texto || 'Ver más'),
              url: String(b.url || 'https://acofi.edu.co'),
              estilo: (b.estilo as EstiloBoton) || 'solido',
              colorBoton: b.colorBoton || edicionActiva?.color_principal || '#1d4ed8',
              colorTexto: b.colorTexto || '#ffffff',
              colorIluminacion: b.colorIluminacion || '#60a5fa',
              alineacion: b.alineacion || 'centro',
              abrirNuevaPestana: b.abrirNuevaPestana !== false,
              colorFondo: b.colorFondo || 'transparent'
            };
          }
          return {
            id: b.id,
            tipo: 'texto',
            contenido: String(b.contenido || ''),
            colorFondo: 'transparent'
          };
        });

        return { ...pag, bloques: bloquesSanitizados };
      });

      let nuevasPaginasTotales: Pagina[] = [];
      let nuevoIndiceIr = indicePaginaActual;

      if (modoInsercionJson === 'reemplazar') {
        nuevasPaginasTotales = paginasListas;
        nuevoIndiceIr = 0;
      } else if (modoInsercionJson === 'despues_actual') {
        nuevasPaginasTotales = [...paginas];
        nuevasPaginasTotales.splice(indicePaginaActual + 1, 0, ...paginasListas);
        nuevoIndiceIr = indicePaginaActual + 1;
      } else {
        nuevasPaginasTotales = [...paginas, ...paginasListas];
        nuevoIndiceIr = paginas.length;
      }

      autoGuardar(nuevasPaginasTotales);
      cambiarPagina(nuevoIndiceIr);
      setJsonLoteInput('');
      setModalImportarJsonAbierto(false);
      mostrarMensaje(`¡${paginasListas.length} página(s) inyectadas y guardadas con éxito!`, 'exito');
    } catch (err: any) {
      mostrarMensaje(`Error en el formato JSON: ${err.message}`, 'error');
    }
  };

  const actualizarConfiguracion = (campo: keyof ConfiguracionEdicion, valor: any) => {
    const nueva = { ...configuracion, [campo]: valor };
    setConfiguracion(nueva);
    autoGuardar(paginas, nueva);
  };

  const actualizarColorTema = async (colorHex: string) => {
    if (!edicionActiva) return;
    setEdicionActiva({ ...edicionActiva, color_principal: colorHex });
    try {
      await supabase.from('ediciones').update({ color_principal: colorHex }).eq('id', edicionActiva.id);
    } catch (err) {
      console.error("Error al actualizar color tema:", err);
    }
  };

  const actualizarEstadoEdicion = async (nuevoEstado: 'publicado' | 'borrador') => {
    if (!edicionActiva) return;
    setEdicionActiva({ ...edicionActiva, estado: nuevoEstado });
    try {
      await supabase.from('ediciones').update({ estado: nuevoEstado }).eq('id', edicionActiva.id);
      setEdiciones(prev => prev.map(e => e.id === edicionActiva.id ? { ...e, estado: nuevoEstado } : e));
      mostrarMensaje(`Estado cambiado a: ${nuevoEstado.toUpperCase()}`, 'exito');
    } catch (err: any) {
      mostrarMensaje(`Error al cambiar estado: ${err.message}`, 'error');
    }
  };

  const manejarGuardadoManual = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!edicionActiva) return;
    setGuardando(true);
    try {
      const jsonToSave = JSON.stringify({ paginas, configuracion });
      await supabase.from('ediciones').update({
        titulo: edicionActiva.titulo, anio: edicionActiva.anio, color_principal: edicionActiva.color_principal,
        estado: edicionActiva.estado, contenido_presentacion: jsonToSave
      }).eq('id', edicionActiva.id);
      mostrarMensaje('Configuración guardada exitosamente.', 'exito');
      await cargarEdiciones();
    } catch (err: any) {
      mostrarMensaje(`Error: ${err.message}`, 'error');
    } finally {
      setGuardando(false);
    }
  };

  const manejarSubidaImagen = async (e: React.ChangeEvent<HTMLInputElement>, bloqueId: string) => {
    const file = e.target.files?.[0];
    if (!file || !edicionActiva) return;

    setSubiendoImagen(prev => ({ ...prev, [bloqueId]: true }));

    try {
      const fileExt = file.name.split('.').pop();
      const fileName = `${Date.now()}-${generarId()}.${fileExt}`;
      const filePath = `edicion_${edicionActiva.anio}/${fileName}`; 

      const { error: uploadError } = await supabase.storage.from('imagenes').upload(filePath, file, { cacheControl: '3600', upsert: false });
      if (uploadError) throw uploadError;

      const { data } = supabase.storage.from('imagenes').getPublicUrl(filePath);

      const nuevasPaginas = [...paginas];
      const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueImagen;
      if (bloque) {
        bloque.url = data.publicUrl;
        autoGuardar(nuevasPaginas);
      }
      mostrarMensaje('Imagen subida con éxito', 'exito');
    } catch (err: any) {
      mostrarMensaje(`Error al subir: ${err.message}`, 'error');
    } finally {
      setSubiendoImagen(prev => ({ ...prev, [bloqueId]: false }));
    }
  };

  // Subida de imagen para firmas individuales dentro de un bloque de firma
  const manejarSubidaImagenFirma = async (e: React.ChangeEvent<HTMLInputElement>, bloqueId: string, firmaIdx: number) => {
    const file = e.target.files?.[0];
    if (!file || !edicionActiva) return;

    const keyLoader = `${bloqueId}_firma_${firmaIdx}`;
    setSubiendoImagen(prev => ({ ...prev, [keyLoader]: true }));

    try {
      const fileExt = file.name.split('.').pop();
      const fileName = `${Date.now()}-${generarId()}.${fileExt}`;
      const filePath = `edicion_${edicionActiva.anio}/firmas/${fileName}`;

      const { error: uploadError } = await supabase.storage.from('imagenes').upload(filePath, file, { cacheControl: '3600', upsert: false });
      if (uploadError) throw uploadError;

      const { data } = supabase.storage.from('imagenes').getPublicUrl(filePath);

      const nuevasPaginas = [...paginas];
      const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueFirma;
      if (bloque && bloque.tipo === 'firma' && bloque.firmas[firmaIdx]) {
        bloque.firmas[firmaIdx].url = data.publicUrl;
        autoGuardar(nuevasPaginas);
      }
      mostrarMensaje('Firma subida con éxito', 'exito');
    } catch (err: any) {
      mostrarMensaje(`Error al subir firma: ${err.message}`, 'error');
    } finally {
      setSubiendoImagen(prev => ({ ...prev, [keyLoader]: false }));
    }
  };

  // Eliminar únicamente la imagen de una firma conservando sus textos (Nombre, Cargo, Info adicional)
  const quitarSoloImagenFirma = (bloqueId: string, firmaIdx: number) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueFirma;
    if (bloque && bloque.tipo === 'firma' && bloque.firmas[firmaIdx]) {
      bloque.firmas[firmaIdx].url = '';
      autoGuardar(nuevasPaginas);
      mostrarMensaje('Imagen de firma eliminada (se conservaron los datos del firmante).', 'exito');
    }
  };

  // Subida e inyección directa de imagen dentro de un bloque de texto (en la posición del cursor o al final)
  const manejarSubidaImagenEnTexto = async (
    e: React.ChangeEvent<HTMLInputElement>,
    bloqueId: string,
    columnaDestino: 1 | 2 | 3,
    anchoValor: string,
    alineacionImg: string,
    bordeImg: 'cuadrado' | 'redondeado' | 'circular' = 'cuadrado'
  ) => {
    const file = e.target.files?.[0];
    if (!file || !edicionActiva) return;

    const keyLoader = `txt_img_${bloqueId}`;
    setSubiendoImagen(prev => ({ ...prev, [keyLoader]: true }));

    try {
      const fileExt = file.name.split('.').pop();
      const fileName = `${Date.now()}-${generarId()}.${fileExt}`;
      const filePath = `edicion_${edicionActiva.anio}/inline/${fileName}`;

      const { error: uploadError } = await supabase.storage.from('imagenes').upload(filePath, file, { cacheControl: '3600', upsert: false });
      if (uploadError) throw uploadError;

      const { data } = supabase.storage.from('imagenes').getPublicUrl(filePath);

      insertarHtmlImagenEnTexto(bloqueId, data.publicUrl, columnaDestino, anchoValor, alineacionImg, bordeImg);
      mostrarMensaje('Imagen insertada en la posición del cursor', 'exito');
    } catch (err: any) {
      mostrarMensaje(`Error al subir imagen: ${err.message}`, 'error');
    } finally {
      setSubiendoImagen(prev => ({ ...prev, [keyLoader]: false }));
      e.target.value = '';
    }
  };

  const construirEstiloImagenTexto = (anchoValor: string, alineacionImg: string, bordeImg: string) => {
    const anchoLimpio = anchoValor.trim() || '50%';
    const radiusVal = bordeImg === 'circular' ? '9999px' : bordeImg === 'redondeado' ? '14px' : '0px';
    let marginStyle = 'margin: 12px auto; display: block;';
    if (alineacionImg === 'izquierda') marginStyle = 'margin: 12px auto 12px 0; display: block;';
    if (alineacionImg === 'derecha') marginStyle = 'margin: 12px 0 12px auto; display: block;';
    return `width: ${anchoLimpio}; max-width: 100%; height: auto; border-radius: ${radiusVal}; ${marginStyle}`;
  };

  const insertarHtmlImagenEnTexto = (
    bloqueId: string,
    urlImg: string,
    columnaDestino: 1 | 2 | 3,
    anchoValor: string,
    alineacionImg: string,
    bordeImg: 'cuadrado' | 'redondeado' | 'circular' = 'cuadrado'
  ) => {
    if (!urlImg.trim()) return;
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTexto;
    if (!bloque || bloque.tipo !== 'texto') return;

    const anchoLimpio = anchoValor.trim() || '50%';
    const claseAlignQuill = alineacionImg === 'centro' ? ' class="ql-align-center"' : alineacionImg === 'derecha' ? ' class="ql-align-right"' : '';
    const imgStyle = construirEstiloImagenTexto(anchoLimpio, alineacionImg, bordeImg);
    const htmlImg = `<p${claseAlignQuill}><img src="${urlImg}" alt="Imagen" width="${anchoLimpio}" data-borde="${bordeImg}" data-alineacion="${alineacionImg}" style="${imgStyle}" /></p>`;

    // Si el usuario dejó el cursor en un punto específico del editor de este bloque, insertamos exactamente ahí
    if (
      ultimoRangoCursorRef.current &&
      ultimoRangoCursorRef.current.bloqueId === bloqueId &&
      ultimoRangoCursorRef.current.columna === columnaDestino &&
      ultimoRangoCursorRef.current.contenedor
    ) {
      try {
        const { range, contenedor } = ultimoRangoCursorRef.current;
        if (contenedor.contains(range.commonAncestorContainer)) {
          const imgElem = document.createElement('img');
          imgElem.setAttribute('src', urlImg);
          imgElem.setAttribute('alt', 'Imagen');
          imgElem.setAttribute('width', anchoLimpio);
          imgElem.setAttribute('data-borde', bordeImg);
          imgElem.setAttribute('data-alineacion', alineacionImg);
          imgElem.setAttribute('style', imgStyle);

          range.insertNode(imgElem);
          const campoCursor = columnaDestino === 2 ? 'contenidoCol2' : columnaDestino === 3 ? 'contenidoCol3' : 'contenido';
          (bloque as any)[campoCursor] = contenedor.innerHTML;
          autoGuardar(nuevasPaginas);
          return;
        }
      } catch (e) {
        // Fallback al método HTML si el rango expiró
      }
    }

    if (columnaDestino === 2) {
      bloque.contenidoCol2 = `${bloque.contenidoCol2 || ''}${htmlImg}`;
    } else if (columnaDestino === 3) {
      bloque.contenidoCol3 = `${bloque.contenidoCol3 || ''}${htmlImg}`;
    } else {
      bloque.contenido = `${bloque.contenido || ''}${htmlImg}`;
    }
    autoGuardar(nuevasPaginas);
  };

  // EXTRAER LISTA DE IMÁGENES ACTUALES DENTRO DE UN BLOQUE DE TEXTO PARA PODER EDITARLAS, MOVERLAS O BORRARLAS INDIVIDUALMENTE
  const extraerImagenesDeColumnaTexto = (htmlColumna?: string) => {
    if (!htmlColumna || !htmlColumna.includes('<img')) return [];
    const regexImg = /<img\b[^>]*>/gi;
    const matches = htmlColumna.match(regexImg) || [];
    return matches.map((imgTag, idx) => {
      const srcMatch = imgTag.match(/src="([^"]*)"/i);
      const widthMatch = imgTag.match(/width="([^"]*)"/i) || imgTag.match(/width:\s*([^;"]+)/i);
      const bordeMatch = imgTag.match(/data-borde="([^"]*)"/i);
      const alignMatch = imgTag.match(/data-alineacion="([^"]*)"/i);

      let bordeDetectado = bordeMatch ? bordeMatch[1] : 'cuadrado';
      if (!bordeMatch && imgTag.includes('9999px')) bordeDetectado = 'circular';
      else if (!bordeMatch && (imgTag.includes('12px') || imgTag.includes('14px'))) bordeDetectado = 'redondeado';

      let alignDetectado = alignMatch ? alignMatch[1] : 'centro';
      if (!alignMatch && imgTag.includes('margin: 12px auto 12px 0')) alignDetectado = 'izquierda';
      else if (!alignMatch && imgTag.includes('margin: 12px 0 12px auto')) alignDetectado = 'derecha';

      return {
        indice: idx,
        tagOriginal: imgTag,
        src: srcMatch ? srcMatch[1] : '',
        ancho: widthMatch ? widthMatch[1].trim() : '50%',
        borde: bordeDetectado as 'cuadrado' | 'redondeado' | 'circular',
        alineacion: alignDetectado as 'izquierda' | 'centro' | 'derecha'
      };
    });
  };

  // MODIFICAR PROPIEDADES (TAMAÑO, ALINEACIÓN, ESQUINAS REDONDAS/CUADRADAS) DE UNA IMAGEN ESPECÍFICA DENTRO DEL TEXTO
  const actualizarImagenEspecificaEnTexto = (
    bloqueId: string,
    columnaDestino: 1 | 2 | 3,
    indiceImg: number,
    nuevosValores: { ancho?: string; alineacion?: 'izquierda' | 'centro' | 'derecha'; borde?: 'cuadrado' | 'redondeado' | 'circular' }
  ) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTexto;
    if (!bloque || bloque.tipo !== 'texto') return;

    const campo = columnaDestino === 2 ? 'contenidoCol2' : columnaDestino === 3 ? 'contenidoCol3' : 'contenido';
    const htmlActual = String((bloque as any)[campo] || '');
    const lista = extraerImagenesDeColumnaTexto(htmlActual);
    const objetivo = lista[indiceImg];
    if (!objetivo) return;

    const nuevoAncho = nuevosValores.ancho ?? objetivo.ancho;
    const nuevaAlineacion = nuevosValores.alineacion ?? objetivo.alineacion;
    const nuevoBorde = nuevosValores.borde ?? objetivo.borde;
    const nuevoEstilo = construirEstiloImagenTexto(nuevoAncho, nuevaAlineacion, nuevoBorde);

    let contador = -1;
    const htmlNuevo = htmlActual.replace(/<img\b[^>]*>/gi, (match) => {
      contador++;
      if (contador !== indiceImg) return match;
      return `<img src="${objetivo.src}" alt="Imagen" width="${nuevoAncho}" data-borde="${nuevoBorde}" data-alineacion="${nuevaAlineacion}" style="${nuevoEstilo}" />`;
    });

    (bloque as any)[campo] = htmlNuevo;
    autoGuardar(nuevasPaginas);
  };

  // MOVER UNA IMAGEN ARRIBA O ABAJO ENTRE LOS PÁRRAFOS DEL BLOQUE DE TEXTO
  const moverImagenEntreParrafosTexto = (
    bloqueId: string,
    columnaDestino: 1 | 2 | 3,
    indiceImg: number,
    direccion: 'arriba' | 'abajo'
  ) => {
    if (typeof document === 'undefined') return;
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTexto;
    if (!bloque || bloque.tipo !== 'texto') return;

    const campo = columnaDestino === 2 ? 'contenidoCol2' : columnaDestino === 3 ? 'contenidoCol3' : 'contenido';
    const htmlActual = String((bloque as any)[campo] || '');

    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = htmlActual;
    const todasLasImgs = tempDiv.querySelectorAll('img');
    const imgTarget = todasLasImgs[indiceImg];
    if (!imgTarget) return;

    // Encontrar el nodo de bloque superior (p, div, h1..h6, li) que contiene la imagen
    let bloquePadre: HTMLElement | null = imgTarget as HTMLElement;
    while (bloquePadre && bloquePadre.parentElement && bloquePadre.parentElement !== tempDiv) {
      bloquePadre = bloquePadre.parentElement;
    }
    if (!bloquePadre) return;

    // Si el párrafo tiene texto además de la imagen, separamos la imagen en su propio párrafo para moverla libremente
    if ((bloquePadre.textContent || '').trim().length > 0) {
      const nuevoParrafoImg = document.createElement('p');
      nuevoParrafoImg.appendChild(imgTarget.cloneNode(true));
      imgTarget.remove();
      if (direccion === 'arriba') {
        tempDiv.insertBefore(nuevoParrafoImg, bloquePadre);
      } else {
        if (bloquePadre.nextSibling) {
          tempDiv.insertBefore(nuevoParrafoImg, bloquePadre.nextSibling);
        } else {
          tempDiv.appendChild(nuevoParrafoImg);
        }
      }
    } else {
      if (direccion === 'arriba' && bloquePadre.previousElementSibling) {
        tempDiv.insertBefore(bloquePadre, bloquePadre.previousElementSibling);
      } else if (direccion === 'abajo' && bloquePadre.nextElementSibling) {
        tempDiv.insertBefore(bloquePadre.nextElementSibling, bloquePadre);
      } else {
        mostrarMensaje(direccion === 'arriba' ? 'La imagen ya está al inicio del bloque.' : 'La imagen ya está al final del bloque.', 'error');
        return;
      }
    }

    (bloque as any)[campo] = tempDiv.innerHTML;
    autoGuardar(nuevasPaginas);
    mostrarMensaje(direccion === 'arriba' ? 'Imagen subida un párrafo' : 'Imagen bajada un párrafo', 'exito');
  };

  // ELIMINAR UNA IMAGEN ESPECÍFICA DENTRO DE UN BLOQUE DE TEXTO
  const eliminarImagenDeColumnaTexto = (bloqueId: string, columnaDestino: 1 | 2 | 3, indiceImg: number) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTexto;
    if (!bloque || bloque.tipo !== 'texto') return;

    const campo = columnaDestino === 2 ? 'contenidoCol2' : columnaDestino === 3 ? 'contenidoCol3' : 'contenido';
    const htmlActual = String((bloque as any)[campo] || '');

    let contador = -1;
    const htmlSinImg = htmlActual
      .replace(/<img\b[^>]*>/gi, (match) => {
        contador++;
        return contador === indiceImg ? '' : match;
      })
      .replace(/<p[^>]*>\s*<\/p>/gi, '');

    (bloque as any)[campo] = htmlSinImg;
    autoGuardar(nuevasPaginas);
    mostrarMensaje('Imagen eliminada del bloque de texto', 'exito');
  };

  const manejarSubidaImagenConfig = async (e: React.ChangeEvent<HTMLInputElement>, campo: keyof ConfiguracionEdicion) => {
    const file = e.target.files?.[0];
    if (!file || !edicionActiva) return;

    setSubiendoImagen(prev => ({ ...prev, [campo]: true }));

    try {
      const fileExt = file.name.split('.').pop();
      const fileName = `${Date.now()}-${generarId()}.${fileExt}`;
      const filePath = `edicion_${edicionActiva.anio}/config/${fileName}`; 

      const { error: uploadError } = await supabase.storage.from('imagenes').upload(filePath, file, { cacheControl: '3600', upsert: false });
      if (uploadError) throw uploadError;

      const { data } = supabase.storage.from('imagenes').getPublicUrl(filePath);

      actualizarConfiguracion(campo, data.publicUrl);
      mostrarMensaje('Recurso visual actualizado', 'exito');
    } catch (err: any) {
      mostrarMensaje(`Error al subir: ${err.message}`, 'error');
    } finally {
      setSubiendoImagen(prev => ({ ...prev, [campo]: false }));
    }
  };

  const moverPagina = (direccion: 'izq' | 'der') => {
    const nuevasPaginas = [...paginas];
    if (direccion === 'izq' && indicePaginaActual > 0) {
      const temp = nuevasPaginas[indicePaginaActual];
      nuevasPaginas[indicePaginaActual] = nuevasPaginas[indicePaginaActual - 1];
      nuevasPaginas[indicePaginaActual - 1] = temp;
      autoGuardar(nuevasPaginas);
      cambiarPagina(indicePaginaActual - 1);
    } else if (direccion === 'der' && indicePaginaActual < paginas.length - 1) {
      const temp = nuevasPaginas[indicePaginaActual];
      nuevasPaginas[indicePaginaActual] = nuevasPaginas[indicePaginaActual + 1];
      nuevasPaginas[indicePaginaActual + 1] = temp;
      autoGuardar(nuevasPaginas);
      cambiarPagina(indicePaginaActual + 1);
    }
  };

  const moverBloque = (bloqueIndex: number, direccion: 'arriba' | 'abajo') => {
    const nuevasPaginas = [...paginas];
    const bloques = nuevasPaginas[indicePaginaActual].bloques;

    if (direccion === 'arriba' && bloqueIndex > 0) {
      const temp = bloques[bloqueIndex];
      bloques[bloqueIndex] = bloques[bloqueIndex - 1];
      bloques[bloqueIndex - 1] = temp;
      autoGuardar(nuevasPaginas);
    } else if (direccion === 'abajo' && bloqueIndex < bloques.length - 1) {
      const temp = bloques[bloqueIndex];
      bloques[bloqueIndex] = bloques[bloqueIndex + 1];
      bloques[bloqueIndex + 1] = temp;
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarCampoPagina = (campo: keyof Pagina, valor: any) => {
    const nuevasPaginas = [...paginas];
    (nuevasPaginas[indicePaginaActual] as any)[campo] = valor;
    autoGuardar(nuevasPaginas);
  };

  const actualizarTituloPagina = (nuevoTitulo: string) => {
    actualizarCampoPagina('titulo', nuevoTitulo);
  };

  const actualizarColorFondoPagina = (color: string) => {
    actualizarCampoPagina('colorFondo', color);
  };

  const agregarPagina = () => {
    const nuevaPagina: Pagina = { 
      id: generarId(), 
      titulo: `Página ${paginas.length + 1}`, 
      subtitulo: '',
      mostrarNumero: true,
      posicion: 'despues_portada',
      bloques: [], 
      colorFondo: '#fdfbf7' 
    };
    const nuevasPaginas = [...paginas];
    nuevasPaginas.splice(indicePaginaActual + 1, 0, nuevaPagina);
    autoGuardar(nuevasPaginas);
    cambiarPagina(indicePaginaActual + 1);
  };

  const duplicarPaginaActual = () => {
    const actual = paginas[indicePaginaActual];
    if (!actual) return;

    const mapaTablasClonadas: Record<string, string> = {};
    const bloquesClonados: Bloque[] = actual.bloques.map((b) => {
      const nuevoId = generarId();
      if (b.tipo === 'tabla') {
        mapaTablasClonadas[b.id] = nuevoId;
        return {
          ...b,
          id: nuevoId,
          nombre: `${b.nombre} (Copia)`,
          data: b.data.map(f => [...f]),
          coloresFilas: b.coloresFilas ? { ...b.coloresFilas } : {},
          coloresColumnas: b.coloresColumnas ? { ...b.coloresColumnas } : {},
          anchosColumnas: b.anchosColumnas ? { ...b.anchosColumnas } : {},
          altosFilas: b.altosFilas ? { ...b.altosFilas } : {}
        };
      }
      if (b.tipo === 'firma') {
        return {
          ...b,
          id: nuevoId,
          firmas: b.firmas.map(f => ({ ...f, id: generarId() }))
        };
      }
      if (b.tipo === 'grafica') {
        return {
          ...b,
          id: nuevoId,
          ejesY: [...(b.ejesY || [])],
          coloresCustom: [...(b.coloresCustom || [])]
        };
      }
      return { ...b, id: nuevoId };
    }).map((b) => {
      if (b.tipo === 'grafica' && mapaTablasClonadas[b.tablaRef]) {
        return { ...b, tablaRef: mapaTablasClonadas[b.tablaRef] };
      }
      return b;
    });

    const paginaClonada: Pagina = {
      ...actual,
      id: generarId(),
      titulo: actual.titulo ? `${actual.titulo} (Copia)` : '',
      bloques: bloquesClonados
    };

    const nuevasPaginas = [...paginas];
    nuevasPaginas.splice(indicePaginaActual + 1, 0, paginaClonada);
    autoGuardar(nuevasPaginas);
    cambiarPagina(indicePaginaActual + 1);
    mostrarMensaje('Página duplicada con éxito.', 'exito');
  };

  const intentarEliminarPagina = () => {
    if (paginas.length <= 1) {
      mostrarMensaje('No puedes eliminar la única página del libro.', 'error');
      return;
    }
    setModalConfirmacion({
      abierto: true, titulo: 'Eliminar Página', mensaje: `¿Estás seguro de que deseas eliminar esta página? Se borrarán todos los bloques y datos que contiene.`, textoBotonConfirmar: 'Eliminar Página', esPeligro: true,
      accionConfirmar: () => {
        const nuevasPaginas = paginas.filter((_, i) => i !== indicePaginaActual);
        autoGuardar(nuevasPaginas);
        cambiarPagina(Math.max(0, indicePaginaActual - 1));
        setModalConfirmacion(prev => ({ ...prev, abierto: false }));
        mostrarMensaje('Página eliminada con éxito', 'exito');
      }
    });
  };

  const agregarBloqueTexto = () => {
    const nuevasPaginas = [...paginas];
    nuevasPaginas[indicePaginaActual].bloques.push({ 
      id: generarId(), 
      tipo: 'texto', 
      contenido: '', 
      columnas: 1,
      contenidoCol2: '',
      contenidoCol3: '',
      colorFondo: 'transparent' 
    });
    autoGuardar(nuevasPaginas);
  };

  const agregarBloqueFirma = () => {
    const nuevasPaginas = [...paginas];
    nuevasPaginas[indicePaginaActual].bloques.push({
      id: generarId(),
      tipo: 'firma',
      alineacion: 'espaciado',
      espacioGeneralSuperior: 0,
      espacioGeneralInferior: 0,
      colorFondo: 'transparent',
      firmas: [
        {
          id: generarId(),
          url: '',
          nombre: 'Luis Alberto González Araujo',
          cargo: 'Director ejecutivo ACOFI',
          infoAdicional: '',
          tamanoFirma: 72,
          espacioSuperior: 0
        }
      ]
    });
    autoGuardar(nuevasPaginas);
  };

  const agregarBloqueBoton = () => {
    const nuevasPaginas = [...paginas];
    nuevasPaginas[indicePaginaActual].bloques.push({
      id: generarId(),
      tipo: 'boton',
      texto: 'Haz clic aquí',
      url: 'https://acofi.edu.co',
      estilo: '4d',
      colorBoton: edicionActiva?.color_principal || '#1d4ed8',
      colorTexto: '#ffffff',
      colorIluminacion: '#60a5fa',
      alineacion: 'centro',
      abrirNuevaPestana: true,
      colorFondo: 'transparent'
    });
    autoGuardar(nuevasPaginas);
  };

  const crearTablaEnBlanco = () => {
    const nuevasPaginas = [...paginas];
    nuevasPaginas[indicePaginaActual].bloques.push({ 
      id: generarId(), 
      tipo: 'tabla', 
      nombre: 'Nueva Tabla de Datos', 
      data: [['Periodo', 'Total', 'Oficial', 'Privada'], ['2024-1', '0', '0', '0']], 
      oculto: false, 
      colorFondo: '#ffffff',
      escala: 100,
      coloresFilas: {},
      coloresColumnas: {},
      paddingFila: 6,
      tamanoTextoTabla: 14,
      anchosColumnas: {},
      altosFilas: {},
      pieBloque: ''
    });
    autoGuardar(nuevasPaginas);
    setModalNuevaTablaAbierto(false);
    mostrarMensaje('Tabla en blanco insertada', 'exito');
  };

  const clonarTablaExistente = (tablaOrigen: BloqueTabla) => {
    const nuevasPaginas = [...paginas];
    const copiaData = tablaOrigen.data.map(fila => [...fila]);
    const nuevaTabla: BloqueTabla = {
      id: generarId(),
      tipo: 'tabla',
      nombre: `${tablaOrigen.nombre || 'Tabla'} (Copia)`,
      data: copiaData,
      oculto: false,
      colorFondo: tablaOrigen.colorFondo || '#ffffff',
      escala: tablaOrigen.escala || 100,
      coloresFilas: tablaOrigen.coloresFilas ? { ...tablaOrigen.coloresFilas } : {},
      coloresColumnas: tablaOrigen.coloresColumnas ? { ...tablaOrigen.coloresColumnas } : {},
      paddingFila: tablaOrigen.paddingFila ?? 6,
      tamanoTextoTabla: tablaOrigen.tamanoTextoTabla ?? 14,
      anchosColumnas: tablaOrigen.anchosColumnas ? { ...tablaOrigen.anchosColumnas } : {},
      altosFilas: tablaOrigen.altosFilas ? { ...tablaOrigen.altosFilas } : {},
      pieBloque: tablaOrigen.pieBloque || ''
    };
    nuevasPaginas[indicePaginaActual].bloques.push(nuevaTabla);
    autoGuardar(nuevasPaginas);
    setModalNuevaTablaAbierto(false);
    mostrarMensaje(`Tabla "${tablaOrigen.nombre}" clonada con éxito`, 'exito');
  };

  const agregarBloqueTabla = () => {
    setModalNuevaTablaAbierto(true);
  };

  const agregarBloqueGrafica = () => {
    const nuevasPaginas = [...paginas];
    nuevasPaginas[indicePaginaActual].bloques.push({ 
      id: generarId(), 
      tipo: 'grafica', 
      nombre: 'Nueva Gráfica', 
      tablaRef: '', 
      tipoGrafica: 'barras', 
      ejeX: 0, 
      ejesY: [], 
      coloresCustom: [], 
      mostrarValores: false, 
      colorFondo: '#ffffff',
      escala: 100,
      pieBloque: ''
    });
    autoGuardar(nuevasPaginas);
  };

  const agregarBloqueImagen = () => {
    const nuevasPaginas = [...paginas];
    nuevasPaginas[indicePaginaActual].bloques.push({ id: generarId(), tipo: 'imagen', url: '', alt: '', pantallaCompleta: false, colorFondo: 'transparent' });
    autoGuardar(nuevasPaginas);
  };

  const intentarEliminarBloque = (bloqueId: string, tipo: string) => {
    setModalConfirmacion({
      abierto: true, titulo: 'Eliminar Bloque', mensaje: `¿Estás seguro de que deseas eliminar este bloque de ${tipo}? Perderás la información que contiene.`, textoBotonConfirmar: 'Eliminar Bloque', esPeligro: true,
      accionConfirmar: () => {
        const nuevasPaginas = [...paginas];
        nuevasPaginas[indicePaginaActual].bloques = nuevasPaginas[indicePaginaActual].bloques.filter(b => b.id !== bloqueId);
        autoGuardar(nuevasPaginas);
        setModalConfirmacion(prev => ({ ...prev, abierto: false }));
      }
    });
  };

  const actualizarBloqueTexto = (bloqueId: string, nuevoContenido: string) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId);
    if (bloque && bloque.tipo === 'texto') { bloque.contenido = nuevoContenido; autoGuardar(nuevasPaginas); }
  };

  const actualizarCampoBloqueTexto = (bloqueId: string, campo: keyof BloqueTexto, valor: any) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTexto;
    if (bloque && bloque.tipo === 'texto') {
      (bloque as any)[campo] = valor;
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarItemFirma = (bloqueId: string, firmaIdx: number, campo: keyof ItemFirma, valor: any) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueFirma;
    if (bloque && bloque.tipo === 'firma' && bloque.firmas[firmaIdx]) {
      (bloque.firmas[firmaIdx] as any)[campo] = valor;
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarPropiedadBloqueFirma = (bloqueId: string, campo: keyof BloqueFirma, valor: any) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueFirma;
    if (bloque && bloque.tipo === 'firma') {
      (bloque as any)[campo] = valor;
      autoGuardar(nuevasPaginas);
    }
  };

  const agregarSegundaFirmaAlBloque = (bloqueId: string) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueFirma;
    if (bloque && bloque.tipo === 'firma') {
      bloque.firmas.push({
        id: generarId(),
        url: '',
        nombre: 'José Miguel Solano Araujo',
        cargo: 'Asistente de proyectos ACOFI',
        infoAdicional: 'Editor de la publicación',
        tamanoFirma: 72,
        espacioSuperior: 0
      });
      autoGuardar(nuevasPaginas);
    }
  };

  const eliminarFirmaDelBloque = (bloqueId: string, firmaIdx: number) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueFirma;
    if (bloque && bloque.tipo === 'firma' && bloque.firmas.length > 1) {
      bloque.firmas = bloque.firmas.filter((_, i) => i !== firmaIdx);
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarAlineacionBloqueFirma = (bloqueId: string, alineacion: BloqueFirma['alineacion']) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueFirma;
    if (bloque && bloque.tipo === 'firma') {
      bloque.alineacion = alineacion;
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarBloqueBoton = (bloqueId: string, campo: keyof BloqueBoton, valor: any) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueBoton;
    if (bloque && bloque.tipo === 'boton') {
      (bloque as any)[campo] = valor;
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarCeldaTabla = (bloqueId: string, filaIndex: number, colIndex: number, valor: string) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId);
    if (bloque && bloque.tipo === 'tabla') { bloque.data[filaIndex][colIndex] = valor; autoGuardar(nuevasPaginas); }
  };

  const alternarVisibilidadTabla = (bloqueId: string) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTabla;
    if (bloque && bloque.tipo === 'tabla') { 
      bloque.oculto = !bloque.oculto; 
      autoGuardar(nuevasPaginas); 
    }
  };

  const actualizarColorFilaTabla = (bloqueId: string, filaIndex: number, colorHex: string) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTabla;
    if (bloque && bloque.tipo === 'tabla') {
      bloque.coloresFilas = { ...(bloque.coloresFilas || {}), [filaIndex]: colorHex };
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarColorColumnaTabla = (bloqueId: string, colIndex: number, colorHex: string) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTabla;
    if (bloque && bloque.tipo === 'tabla') {
      bloque.coloresColumnas = { ...(bloque.coloresColumnas || {}), [colIndex]: colorHex };
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarPropiedadTabla = (bloqueId: string, campo: keyof BloqueTabla, valor: any) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTabla;
    if (bloque && bloque.tipo === 'tabla') {
      (bloque as any)[campo] = valor;
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarAnchoColumnaTabla = (bloqueId: string, colIndex: number, anchoPx: number) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTabla;
    if (bloque && bloque.tipo === 'tabla') {
      const nuevosAnchos = { ...(bloque.anchosColumnas || {}) };
      if (!anchoPx || anchoPx <= 0) delete nuevosAnchos[colIndex];
      else nuevosAnchos[colIndex] = anchoPx;
      bloque.anchosColumnas = nuevosAnchos;
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarAltoFilaTabla = (bloqueId: string, filaIndex: number, altoPx: number) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTabla;
    if (bloque && bloque.tipo === 'tabla') {
      const nuevosAltos = { ...(bloque.altosFilas || {}) };
      if (!altoPx || altoPx <= 0) delete nuevosAltos[filaIndex];
      else nuevosAltos[filaIndex] = altoPx;
      bloque.altosFilas = nuevosAltos;
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarEscalaBloque = (bloqueId: string, nuevaEscala: number) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId);
    if (bloque) {
      bloque.escala = nuevaEscala;
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarGrafica = (bloqueId: string, campo: keyof BloqueGrafica, valor: any) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueGrafica;
    if (bloque && bloque.tipo === 'grafica') {
      (bloque as any)[campo] = valor;
      if (campo === 'tablaRef') { bloque.ejeX = 0; bloque.ejesY = []; }
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarGraficaColor = (bloqueId: string, indiceColor: number, colorHex: string) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueGrafica;
    if (bloque && bloque.tipo === 'grafica') {
      const nuevosColores = [...(bloque.coloresCustom || [])];
      nuevosColores[indiceColor] = colorHex;
      bloque.coloresCustom = nuevosColores;
      autoGuardar(nuevasPaginas);
    }
  };

  const actualizarColorFondoBloque = (bloqueId: string, colorHex: string) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId);
    if (bloque) {
      bloque.colorFondo = colorHex;
      autoGuardar(nuevasPaginas);
    }
  };

  const agregarFilaTabla = (bloqueId: string) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId);
    if (bloque && bloque.tipo === 'tabla') {
      const numColumnas = bloque.data[0]?.length || 1;
      bloque.data.push(Array(numColumnas).fill('0'));
      autoGuardar(nuevasPaginas);
    }
  };

  const insertarFilaEnPosicion = (bloqueId: string, despuesDeIndex: number) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId);
    if (bloque && bloque.tipo === 'tabla') {
      const numColumnas = bloque.data[0]?.length || 1;
      bloque.data.splice(despuesDeIndex + 1, 0, Array(numColumnas).fill('0'));
      autoGuardar(nuevasPaginas);
    }
  };

  const moverFilaTabla = (bloqueId: string, filaIndex: number, direccion: 'arriba' | 'abajo') => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTabla;
    if (bloque && bloque.tipo === 'tabla') {
      const destino = direccion === 'arriba' ? filaIndex - 1 : filaIndex + 1;
      if (destino < 0 || destino >= bloque.data.length) return;
      const temp = bloque.data[filaIndex];
      bloque.data[filaIndex] = bloque.data[destino];
      bloque.data[destino] = temp;
      if (bloque.coloresFilas) {
        const cActual = bloque.coloresFilas[filaIndex];
        const cDestino = bloque.coloresFilas[destino];
        const nuevosColores = { ...bloque.coloresFilas };
        if (cDestino !== undefined) nuevosColores[filaIndex] = cDestino; else delete nuevosColores[filaIndex];
        if (cActual !== undefined) nuevosColores[destino] = cActual; else delete nuevosColores[destino];
        bloque.coloresFilas = nuevosColores;
      }
      if (bloque.altosFilas) {
        const aActual = bloque.altosFilas[filaIndex];
        const aDestino = bloque.altosFilas[destino];
        const nuevosAltos = { ...bloque.altosFilas };
        if (aDestino !== undefined) nuevosAltos[filaIndex] = aDestino; else delete nuevosAltos[filaIndex];
        if (aActual !== undefined) nuevosAltos[destino] = aActual; else delete nuevosAltos[destino];
        bloque.altosFilas = nuevosAltos;
      }
      autoGuardar(nuevasPaginas);
    }
  };

  const eliminarFilaTabla = (bloqueId: string, filaIndex: number) => {
    const bloque = paginas[indicePaginaActual]?.bloques.find(b => b.id === bloqueId);
    if (!bloque || bloque.tipo !== 'tabla') return;
    if (bloque.data.length <= 2) {
      mostrarMensaje('La tabla debe conservar al menos la cabecera y una fila de datos.', 'error');
      return;
    }
    const nombreFila = bloque.data[filaIndex]?.[0] || `Fila ${filaIndex + 1}`;
    setModalConfirmacion({
      abierto: true,
      titulo: 'Eliminar Fila de Tabla',
      mensaje: `¿Estás seguro de que deseas eliminar la fila "${nombreFila}"? Esta acción no se puede deshacer.`,
      textoBotonConfirmar: 'Sí, Eliminar Fila',
      esPeligro: true,
      accionConfirmar: () => {
        const nuevasPaginas = [...paginas];
        const b = nuevasPaginas[indicePaginaActual].bloques.find(bl => bl.id === bloqueId) as BloqueTabla;
        if (b && b.tipo === 'tabla') {
          b.data = b.data.filter((_, idx) => idx !== filaIndex);
          autoGuardar(nuevasPaginas);
        }
        setModalConfirmacion(prev => ({ ...prev, abierto: false }));
      }
    });
  };

  const agregarColumnaTabla = (bloqueId: string) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId);
    if (bloque && bloque.tipo === 'tabla') {
      bloque.data = bloque.data.map((fila, i) => [...fila, i === 0 ? 'Nueva Col' : '0']);
      autoGuardar(nuevasPaginas);
    }
  };

  const insertarColumnaEnPosicion = (bloqueId: string, despuesDeColIndex: number) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId);
    if (bloque && bloque.tipo === 'tabla') {
      bloque.data = bloque.data.map((fila, i) => {
        const copia = [...fila];
        copia.splice(despuesDeColIndex + 1, 0, i === 0 ? 'Nueva Col' : '0');
        return copia;
      });
      autoGuardar(nuevasPaginas);
    }
  };

  const moverColumnaTabla = (bloqueId: string, colIndex: number, direccion: 'izq' | 'der') => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTabla;
    if (bloque && bloque.tipo === 'tabla') {
      const totalCols = bloque.data[0]?.length || 0;
      const destino = direccion === 'izq' ? colIndex - 1 : colIndex + 1;
      if (destino < 0 || destino >= totalCols) return;
      bloque.data = bloque.data.map(fila => {
        const copia = [...fila];
        const temp = copia[colIndex];
        copia[colIndex] = copia[destino];
        copia[destino] = temp;
        return copia;
      });
      if (bloque.coloresColumnas) {
        const cActual = bloque.coloresColumnas[colIndex];
        const cDestino = bloque.coloresColumnas[destino];
        const nuevosColores = { ...bloque.coloresColumnas };
        if (cDestino !== undefined) nuevosColores[colIndex] = cDestino; else delete nuevosColores[colIndex];
        if (cActual !== undefined) nuevosColores[destino] = cActual; else delete nuevosColores[destino];
        bloque.coloresColumnas = nuevosColores;
      }
      if (bloque.anchosColumnas) {
        const wActual = bloque.anchosColumnas[colIndex];
        const wDestino = bloque.anchosColumnas[destino];
        const nuevosAnchos = { ...bloque.anchosColumnas };
        if (wDestino !== undefined) nuevosAnchos[colIndex] = wDestino; else delete nuevosAnchos[colIndex];
        if (wActual !== undefined) nuevosAnchos[destino] = wActual; else delete nuevosAnchos[destino];
        bloque.anchosColumnas = nuevosAnchos;
      }
      autoGuardar(nuevasPaginas);
    }
  };

  const eliminarColumnaTabla = (bloqueId: string, colIndex: number) => {
    const bloque = paginas[indicePaginaActual]?.bloques.find(b => b.id === bloqueId);
    if (!bloque || bloque.tipo !== 'tabla') return;
    if ((bloque.data[0]?.length || 0) <= 2) {
      mostrarMensaje('La tabla debe tener al menos 2 columnas.', 'error');
      return;
    }
    const nombreCol = bloque.data[0]?.[colIndex] || `Columna ${colIndex + 1}`;
    setModalConfirmacion({
      abierto: true,
      titulo: 'Eliminar Columna de Tabla',
      mensaje: `¿Estás seguro de que deseas eliminar la columna "${nombreCol}" y todos sus datos?`,
      textoBotonConfirmar: 'Sí, Eliminar Columna',
      esPeligro: true,
      accionConfirmar: () => {
        const nuevasPaginas = [...paginas];
        const b = nuevasPaginas[indicePaginaActual].bloques.find(bl => bl.id === bloqueId) as BloqueTabla;
        if (b && b.tipo === 'tabla') {
          b.data = b.data.map(fila => fila.filter((_, idx) => idx !== colIndex));
          autoGuardar(nuevasPaginas);
        }
        setModalConfirmacion(prev => ({ ...prev, abierto: false }));
      }
    });
  };
  const intentarDuplicarEdicion = () => {
    if (!edicionActiva) return;
    setModalConfirmacion({
      abierto: true, titulo: 'Duplicar Edición', mensaje: `¿Deseas crear una copia exacta de la edición ${edicionActiva.anio} con todos sus bloques y configuraciones?`, textoBotonConfirmar: 'Sí, Duplicar Edición', esPeligro: false,
      accionConfirmar: async () => {
        setModalConfirmacion(prev => ({ ...prev, abierto: false }));
        setGuardando(true);
        const nuevoAnio = (parseInt(edicionActiva.anio) + 1).toString();
        try {
          const jsonToSave = JSON.stringify({ paginas, configuracion });
          const { error } = await supabase.from('ediciones').insert([{
            titulo: edicionActiva.titulo, anio: nuevoAnio, color_principal: edicionActiva.color_principal, estado: 'borrador', contenido_presentacion: jsonToSave
          }]);
          if (error) throw error;
          mostrarMensaje(`Edición ${nuevoAnio} creada con éxito`, 'exito');
          await cargarEdiciones();
        } catch (err: any) {
          mostrarMensaje(`Error al duplicar: ${err.message}`, 'error');
        } finally {
          setGuardando(false);
        }
      }
    });
  };

  const intentarEliminarEdicion = () => {
    if (!edicionActiva) return;
    setModalConfirmacion({
      abierto: true, titulo: 'Eliminar Edición Permanente', mensaje: `Estás a punto de eliminar PERMANENTEMENTE la edición ${edicionActiva.anio} y todos sus contenidos. Esta acción no se puede deshacer.`, textoBotonConfirmar: 'Eliminar Definitivamente', esPeligro: true,
      accionConfirmar: async () => {
        setModalConfirmacion(prev => ({ ...prev, abierto: false }));
        setEliminando(true);
        try {
          const { error } = await supabase.from('ediciones').delete().eq('id', edicionActiva.id);
          if (error) throw error;
          mostrarMensaje(`La edición ${edicionActiva.anio} fue eliminada del sistema.`, 'exito');
          setEdicionActiva(null);
          await cargarEdiciones();
        } catch (err: any) {
          mostrarMensaje(`Error al eliminar: ${err.message}`, 'error');
        } finally {
          setEliminando(false);
        }
      }
    });
  };

  // CAPTURA INTELIGENTE TANTO DE LA POSICIÓN EXACTA DEL CURSOR COMO DEL TEXTO SOMBREADO EN EL EDITOR
  const capturarSeleccionEditor = (
    bloqueId: string,
    columna: 1 | 2 | 3,
    e: React.MouseEvent<HTMLDivElement> | React.KeyboardEvent<HTMLDivElement>
  ) => {
    if (typeof window === 'undefined') return;
    const sel = window.getSelection();
    if (!sel || sel.rangeCount === 0) return;

    const range = sel.getRangeAt(0);
    const contenedor = e.currentTarget.querySelector('.ql-editor') as HTMLElement | null;
    if (!contenedor || !contenedor.contains(range.commonAncestorContainer)) return;

    // Guardar siempre la última posición del cursor para poder insertar imágenes justo en medio del texto
    ultimoRangoCursorRef.current = {
      bloqueId,
      columna,
      range: range.cloneRange(),
      contenedor
    };

    const selectImgCol = document.getElementById(`img-col-${bloqueId}`) as HTMLSelectElement | null;
    if (selectImgCol) selectImgCol.value = String(columna);

    const texto = sel.toString().trim();
    if (!texto) return;

    rangoSeleccionRef.current = range.cloneRange();
    editorContenedorRef.current = contenedor;
    setSeleccionTextoActiva({ bloqueId, columna, texto });

    const inputTxt = document.getElementById(`link-txt-${bloqueId}`) as HTMLInputElement | null;
    if (inputTxt) inputTxt.value = texto;
    const selectCol = document.getElementById(`link-col-${bloqueId}`) as HTMLSelectElement | null;
    if (selectCol) selectCol.value = String(columna);
  };

  // ENLAZAR EL TEXTO SELECCIONADO (O ESCRITO) CON PERSONALIZACIÓN DE COLOR Y SUBRAYADO (SÍ/NO)
  const enlazarTextoEnColumna = (
    bloqueId: string,
    columnaDestino: 1 | 2 | 3,
    textoManual: string,
    tipoDestino: 'pagina' | 'url',
    destinoValor: string,
    colorEnlace: string = '#1d4ed8',
    conSubrayado: boolean = true
  ) => {
    const nuevasPaginas = [...paginas];
    const bloque = nuevasPaginas[indicePaginaActual].bloques.find(b => b.id === bloqueId) as BloqueTexto;
    if (!bloque || bloque.tipo !== 'texto') return;

    const colFinal = seleccionTextoActiva && seleccionTextoActiva.bloqueId === bloqueId
      ? seleccionTextoActiva.columna
      : columnaDestino;

    const campo = colFinal === 2 ? 'contenidoCol2' : colFinal === 3 ? 'contenidoCol3' : 'contenido';
    const htmlActual = String((bloque as any)[campo] || '');

    const hrefFinal = tipoDestino === 'pagina' ? `#pagina:${destinoValor}` : (destinoValor.trim() || 'https://acofi.edu.co');
    const decoStyle = conSubrayado ? 'underline' : 'none';
    const estiloLink = `color: ${colorEnlace} !important; text-decoration: ${decoStyle} !important; font-weight: 700; cursor: pointer;`;

    if (
      seleccionTextoActiva &&
      seleccionTextoActiva.bloqueId === bloqueId &&
      rangoSeleccionRef.current &&
      editorContenedorRef.current
    ) {
      try {
        const range = rangoSeleccionRef.current;
        const textoSombreado = range.toString() || seleccionTextoActiva.texto;
        const aElem = document.createElement('a');
        aElem.setAttribute('href', hrefFinal);
        if (tipoDestino === 'pagina') {
          aElem.setAttribute('data-internal-page', destinoValor);
        } else {
          aElem.setAttribute('target', '_blank');
          aElem.setAttribute('rel', 'noreferrer');
        }
        aElem.setAttribute('data-custom-link', 'true');
        aElem.setAttribute('style', estiloLink);
        aElem.textContent = textoManual.trim() || textoSombreado;

        range.deleteContents();
        range.insertNode(aElem);

        const nuevoHtmlEditor = editorContenedorRef.current.innerHTML;
        (bloque as any)[campo] = nuevoHtmlEditor;
        autoGuardar(nuevasPaginas);

        setSeleccionTextoActiva(null);
        rangoSeleccionRef.current = null;
        mostrarMensaje(tipoDestino === 'pagina' ? 'Texto seleccionado enlazado a la página' : 'Texto seleccionado enlazado a la URL', 'exito');
        return;
      } catch (err) {
        // Fallback al reemplazo HTML
      }
    }

    const textoSeleccionadoNavegador = typeof window !== 'undefined' ? (window.getSelection()?.toString().trim() || '') : '';
    const textoUsar = textoManual.trim() || seleccionTextoActiva?.texto || textoSeleccionadoNavegador || 'Ir a la página';

    const attrsExtra = tipoDestino === 'pagina'
      ? `data-internal-page="${destinoValor}"`
      : `target="_blank" rel="noreferrer"`;

    const tagEnlace = `<a href="${hrefFinal}" ${attrsExtra} data-custom-link="true" style="${estiloLink}">${textoUsar}</a>`;

    if (textoUsar && htmlActual.includes(textoUsar)) {
      (bloque as any)[campo] = htmlActual.replace(textoUsar, tagEnlace);
    } else {
      (bloque as any)[campo] = `${htmlActual}<p>${tagEnlace}</p>`;
    }

    setSeleccionTextoActiva(null);
    rangoSeleccionRef.current = null;
    autoGuardar(nuevasPaginas);
    mostrarMensaje(tipoDestino === 'pagina' ? 'Texto enlazado a la página seleccionada' : 'Texto enlazado a la URL', 'exito');
  };

  const getBloqueIcon = (b: Bloque) => {
    if (b.tipo === 'texto') return <Type size={14} />;
    if (b.tipo === 'firma') return <PenTool size={14} />;
    if (b.tipo === 'boton') return <MousePointerClick size={14} />;
    if (b.tipo === 'tabla') return <TableIcon size={14} />;
    if (b.tipo === 'imagen') return <ImageIcon size={14} />;
    if (b.tipo === 'grafica') {
      if (b.tipoGrafica === 'mapa_colombia') return <MapIcon size={14} />;
      return <BarChart3 size={14} />;
    }
    return null;
  };

  const RenderUploadConfig = ({ campo, label, placeholder }: { campo: keyof ConfiguracionEdicion, label: string, placeholder: string }) => (
    <div className="bg-slate-50 p-6 rounded-xl border border-slate-200">
      <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-3">{label}</label>
      <div className="flex flex-col gap-4">
        <div className="relative group">
          <input
            type="file"
            accept="image/*"
            onChange={(e) => manejarSubidaImagenConfig(e, campo)}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
            disabled={subiendoImagen[campo]}
          />
          <button type="button" className={`w-full px-4 py-3 rounded-xl text-sm font-bold shadow-sm flex justify-center items-center gap-2 transition-colors ${subiendoImagen[campo] ? 'bg-slate-200 text-slate-500 cursor-not-allowed' : 'bg-white border border-slate-300 text-slate-700 hover:bg-slate-100 group-hover:border-blue-400'}`}>
            {subiendoImagen[campo] ? <RefreshCcw size={16} className="animate-spin" /> : <Upload size={16} />}
            {subiendoImagen[campo] ? 'Subiendo...' : 'Subir Archivo'}
          </button>
        </div>
        <div className="flex items-center justify-center"><span className="text-[10px] font-black text-slate-400 uppercase">O pegar URL</span></div>
        <input
          type="text"
          value={configuracion[campo] as string || ''}
          onChange={(e) => actualizarConfiguracion(campo, e.target.value)}
          className="w-full p-3 bg-white border border-slate-200 rounded-lg text-sm outline-none focus:border-blue-500"
          placeholder={placeholder}
        />
        {(configuracion[campo] as string) && (
          <div className="mt-2 relative rounded overflow-hidden border border-slate-200 bg-white flex justify-center p-2 h-24">
            <img src={configuracion[campo] as string} alt="Previa" className="h-full w-auto object-contain" />
            <button type="button" onClick={() => actualizarConfiguracion(campo, '')} className="absolute top-1 right-1 bg-white rounded-full p-1 shadow hover:text-red-500"><X size={12} /></button>
          </div>
        )}
      </div>
    </div>
  );

  if (revisandoSesion) {
    return <div className="flex h-screen items-center justify-center bg-slate-50 font-sans"><div className="w-12 h-12 border-4 border-blue-600 border-t-transparent rounded-full animate-spin"></div></div>;
  }

  if (!sesion) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4 font-sans">
        <div className="max-w-md w-full bg-white rounded-3xl shadow-xl border border-slate-200 p-8 md:p-10 animate-in fade-in zoom-in-95 duration-500">
          <div className="flex justify-center mb-6">
            <div className="p-4 bg-blue-600 text-white rounded-2xl shadow-lg shadow-blue-600/30">
              <Lock size={32} strokeWidth={2} />
            </div>
          </div>
          <h2 className="text-3xl font-black text-center text-slate-900 mb-2 tracking-tight">Acceso Restringido</h2>
          <p className="text-center text-slate-500 mb-8 font-medium">Panel de administración exclusivo de ACOFI.</p>
          
          <form onSubmit={manejarLogin} className="space-y-5">
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Correo Electrónico</label>
              <input 
                type="email" 
                value={emailLogin} 
                onChange={(e) => setEmailLogin(e.target.value)} 
                required 
                className="w-full px-5 py-4 bg-slate-50 border-2 border-slate-200 rounded-xl text-slate-800 outline-none focus:border-blue-500 transition-colors font-medium"
                placeholder="admin@acofi.edu.co"
              />
            </div>
            <div>
              <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Contraseña</label>
              <input 
                type="password" 
                value={passwordLogin} 
                onChange={(e) => setPasswordLogin(e.target.value)} 
                required 
                className="w-full px-5 py-4 bg-slate-50 border-2 border-slate-200 rounded-xl text-slate-800 outline-none focus:border-blue-500 transition-colors font-medium"
                placeholder="••••••••"
              />
            </div>

            {errorLogin && (
              <div className="p-4 bg-red-50 border border-red-200 text-red-600 rounded-xl text-sm font-bold flex items-center gap-2">
                <AlertCircle size={16} /> {errorLogin}
              </div>
            )}

            <button 
              type="submit" 
              disabled={cargandoLogin} 
              className="w-full bg-blue-600 hover:bg-blue-700 text-white font-black py-4 rounded-xl shadow-lg shadow-blue-600/20 transition-all disabled:opacity-70 flex justify-center items-center gap-2 mt-2"
            >
              {cargandoLogin ? <RefreshCcw className="animate-spin" size={20} /> : 'Iniciar Sesión'}
            </button>
          </form>
        </div>
      </div>
    );
  }

  const paginaActiva = paginas[indicePaginaActual];
  const todasLasTablasAdmin = paginas.flatMap((p, pIdx) => 
    p.bloques.filter(b => b.tipo === 'tabla').map(b => ({ ...(b as BloqueTabla), nombrePagina: p.titulo || `Hoja ${pIdx + 1}` }))
  );

  const parseValor = (val: any) => {
    if (val === undefined || val === null) return 0;
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
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
    return isNaN(n) ? 0 : n;
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

  const esGraficaPorSegmento = (tipo: TipoGrafica) =>
    ['torta', 'torta_rellena', 'torta_3d_rellena', 'rosa_polar', 'dona', 'torta_4d', 'semicirculo', 'barras_radiales', 'embudo', 'treemap'].includes(tipo);

  const formatearDatosParaGrafica = (bloqueGrafica: BloqueGrafica) => {
    const tablaFuente = todasLasTablasAdmin.find(t => t.id === bloqueGrafica.tablaRef);
    if (!tablaFuente || !tablaFuente.data || tablaFuente.data.length < 2) return { datos: [], cabeceras: [], ejesYUsar: [], sufijoPorcentaje: false, columnasPorcentaje: {} as Record<string, boolean> };

    const cabecerasOriginales = tablaFuente.data[0] || [];
    const cabeceras = cabecerasOriginales.map((c, i) => String(c || '').trim() || `Columna_${i + 1}`);
    const filas = tablaFuente.data.slice(1);

    const ejeXSeguro = Number(bloqueGrafica.ejeX) || 0;
    const ejesYUsar = (bloqueGrafica.ejesY && Array.isArray(bloqueGrafica.ejesY) && bloqueGrafica.ejesY.length > 0) 
        ? bloqueGrafica.ejesY.map(Number) 
        : cabeceras.map((_, i) => i).filter(i => i !== ejeXSeguro);

    const indiceMapa = ejesYUsar.length > 0 ? ejesYUsar[0] : 1;

    let detectaPorcentajeGlobal = false;
    const columnasPorcentaje: Record<string, boolean> = {};

    ejesYUsar.forEach(idx => {
      const nombreCab = cabeceras[idx];
      const cabTienePorcentaje = nombreCab.includes('%') || nombreCab.toLowerCase().includes('porcentaje');
      const filaTienePorcentaje = filas.some(f => String(f[idx] || '').includes('%'));
      if (cabTienePorcentaje || filaTienePorcentaje) {
        columnasPorcentaje[nombreCab] = true;
        detectaPorcentajeGlobal = true;
      }
    });

    const datosFormateados = filas.map((fila, fIdx) => {
      const rawVal = String(fila[indiceMapa] || '');
      const nombreCabMapa = cabeceras[indiceMapa] || '';
      const esPorcMapa = rawVal.includes('%') || !!columnasPorcentaje[nombreCabMapa];

      const obj: any = { 
          name: String(fila[ejeXSeguro] || `Ítem ${fIdx + 1}`).trim(),
          valorMapa: parseValor(fila[indiceMapa]),
          esPorcentaje: esPorcMapa,
          _idx: fIdx,
          fill: bloqueGrafica.coloresCustom?.[fIdx] || COLORES_GRAFICAS_DEFAULT[fIdx % COLORES_GRAFICAS_DEFAULT.length]
      };
      ejesYUsar.forEach(idx => {
        const nombreCabecera = cabeceras[idx];
        obj[nombreCabecera] = parseValor(fila[idx]);
      });
      return obj;
    });

    return { datos: datosFormateados, cabeceras, ejesYUsar, sufijoPorcentaje: detectaPorcentajeGlobal, columnasPorcentaje };
  };

  // RENDERIZADOR REAL DE GRÁFICAS PARA LA VISTA PREVIA Y EXPORTACIÓN PDF DEL ADMIN
  const renderizarGrafica = (bloque: BloqueGrafica) => {
    const { datos, cabeceras, ejesYUsar, sufijoPorcentaje, columnasPorcentaje } = formatearDatosParaGrafica(bloque);
    if (datos.length === 0 || ejesYUsar.length === 0) return <div className="p-8 text-center text-stone-400 italic">Faltan datos en la tabla conectada.</div>;

    const nombresEjesY = ejesYUsar.map(idx => cabeceras[idx]);
    const nombreY = nombresEjesY[0];
    const unicaSerie = nombresEjesY.length === 1;

    const formatearEtiqueta = (v: any, nombreSerie?: string) => {
      const num = Number(v);
      if (isNaN(num)) return String(v);
      const llevaPorc = nombreSerie ? (columnasPorcentaje[nombreSerie] || sufijoPorcentaje) : sufijoPorcentaje;
      return `${num.toLocaleString('es-CO', { maximumFractionDigits: 2 })}${llevaPorc ? '%' : ''}`;
    };

    const abrirSelectorDirecto = (indice: number, etiqueta: string) => {
      setSelectorColorActivo({ bloqueId: bloque.id, indice, etiqueta });
    };

    switch (bloque.tipoGrafica) {
      case 'barras':
        return (
          <ResponsiveContainer width="100%" height={360}>
            <BarChart data={datos} margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis domain={[0, 'auto']} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), name]} cursor={{ fill: '#f5f5f4' }} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              {nombresEjesY.map((nombre, i) => (
                <Bar 
                  key={nombre} 
                  dataKey={nombre} 
                  fill={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} 
                  radius={[6, 6, 0, 0]} 
                  isAnimationActive={false}
                  style={{ cursor: 'pointer' }}
                  onClick={(data: any, idx: number) => abrirSelectorDirecto(unicaSerie ? idx : i, unicaSerie ? `${data?.name || `Ítem ${idx + 1}`}` : nombre)}
                >
                  {unicaSerie && datos.map((entry: any, idx: number) => (
                    <Cell 
                      key={`cell-bar-${idx}`} 
                      fill={bloque.coloresCustom?.[idx] || COLORES_GRAFICAS_DEFAULT[idx % COLORES_GRAFICAS_DEFAULT.length]} 
                      cursor="pointer"
                      onClick={() => abrirSelectorDirecto(idx, entry.name)}
                    />
                  ))}
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                </Bar>
              ))}
            </BarChart>
          </ResponsiveContainer>
        );

      case 'barras_horizontales':
        return (
          <ResponsiveContainer width="100%" height={360}>
            <BarChart layout="vertical" data={datos} margin={{ top: 20, right: 55, left: 40, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" horizontal={false} stroke="#e7e5e4" />
              <XAxis type="number" domain={[0, 'auto']} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <YAxis type="category" dataKey="name" tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} width={90} />
              <RechartsTooltip formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), name]} cursor={{ fill: '#f5f5f4' }} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              {nombresEjesY.map((nombre, i) => (
                <Bar 
                  key={nombre} 
                  dataKey={nombre} 
                  fill={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} 
                  radius={[0, 6, 6, 0]} 
                  isAnimationActive={false}
                  style={{ cursor: 'pointer' }}
                  onClick={(data: any, idx: number) => abrirSelectorDirecto(unicaSerie ? idx : i, unicaSerie ? `${data?.name || `Ítem ${idx + 1}`}` : nombre)}
                >
                  {unicaSerie && datos.map((entry: any, idx: number) => (
                    <Cell 
                      key={`cell-hbar-${idx}`} 
                      fill={bloque.coloresCustom?.[idx] || COLORES_GRAFICAS_DEFAULT[idx % COLORES_GRAFICAS_DEFAULT.length]} 
                      cursor="pointer"
                      onClick={() => abrirSelectorDirecto(idx, entry.name)}
                    />
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
          <ResponsiveContainer width="100%" height={360}>
            <BarChart data={datos} margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <defs>
                <filter id={`shadowStack-admin-${bloque.id}`} x="-20%" y="-20%" width="140%" height="140%">
                  <feDropShadow dx="3" dy="5" stdDeviation="3" floodColor="#000000" floodOpacity="0.22" />
                </filter>
              </defs>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis domain={[0, 'auto']} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), name]} cursor={{ fill: '#f5f5f4' }} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              {nombresEjesY.map((nombre, i) => (
                <Bar 
                  key={nombre} 
                  dataKey={nombre} 
                  stackId="a" 
                  fill={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} 
                  radius={i === nombresEjesY.length - 1 ? [8, 8, 0, 0] : [0, 0, 0, 0]} 
                  style={bloque.tipoGrafica === 'barras_3d_apiladas' ? { filter: `url(#shadowStack-admin-${bloque.id})`, cursor: 'pointer' } : { cursor: 'pointer' }}
                  onClick={() => abrirSelectorDirecto(i, nombre)}
                  isAnimationActive={false}
                >
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="inside" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#1e293b', fontSize: 11, fontWeight: 'bold' }} />}
                </Bar>
              ))}
            </BarChart>
          </ResponsiveContainer>
        );

      case 'barras_4d':
        return (
          <ResponsiveContainer width="100%" height={360}>
            <BarChart data={datos} margin={{ top: 30, right: 30, left: 0, bottom: 5 }}>
              <defs>
                <filter id={`shadow4d-admin-${bloque.id}`} x="-20%" y="-20%" width="140%" height="140%">
                  <feDropShadow dx="4" dy="6" stdDeviation="4" floodColor="#000000" floodOpacity="0.25" />
                </filter>
              </defs>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12, fontWeight: 'bold' }} />
              <YAxis domain={[0, 'auto']} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), name]} cursor={{ fill: '#f5f5f4' }} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.15)' }} />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              {nombresEjesY.map((nombre, i) => {
                const baseColor = bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length];
                return (
                  <Bar 
                    key={nombre} 
                    dataKey={nombre} 
                    fill={baseColor} 
                    radius={[12, 12, 4, 4]} 
                    style={{ filter: `url(#shadow4d-admin-${bloque.id})`, cursor: 'pointer' }} 
                    onClick={(data: any, idx: number) => abrirSelectorDirecto(unicaSerie ? idx : i, unicaSerie ? `${data?.name || `Ítem ${idx + 1}`}` : nombre)}
                    isAnimationActive={false}
                  >
                    {unicaSerie && datos.map((entry: any, idx: number) => (
                      <Cell 
                        key={`cell-4dbar-${idx}`} 
                        fill={bloque.coloresCustom?.[idx] || COLORES_GRAFICAS_DEFAULT[idx % COLORES_GRAFICAS_DEFAULT.length]} 
                        cursor="pointer"
                        onClick={() => abrirSelectorDirecto(idx, entry.name)}
                      />
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
          <ResponsiveContainer width="100%" height={360}>
            <PieChart>
              <Pie 
                data={datos} 
                cx="50%" 
                cy={bloque.tipoGrafica === 'semicirculo' ? '72%' : '50%'} 
                startAngle={bloque.tipoGrafica === 'semicirculo' ? 180 : 0}
                endAngle={bloque.tipoGrafica === 'semicirculo' ? 0 : 360}
                innerRadius={bloque.tipoGrafica === 'dona' ? 80 : (bloque.tipoGrafica === 'semicirculo' ? 65 : 0)} 
                outerRadius={125} 
                paddingAngle={bloque.tipoGrafica === 'dona' || bloque.tipoGrafica === 'semicirculo' ? 3 : (bloque.tipoGrafica === 'torta_rellena' ? 0 : 1)} 
                dataKey={nombreY} 
                nameKey="name" 
                label={bloque.mostrarValores ? ({ name, value }) => `${name}: ${formatearEtiqueta(value, nombreY)}` : false}
                isAnimationActive={false}
              >
                {datos.map((entry: any, i: number) => (
                  <Cell 
                    key={`cell-${i}`} 
                    fill={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} 
                    stroke={bloque.tipoGrafica === 'torta_rellena' ? '#ffffff' : 'none'}
                    strokeWidth={bloque.tipoGrafica === 'torta_rellena' ? 1.5 : 0}
                    style={{ cursor: 'pointer' }}
                    onClick={() => abrirSelectorDirecto(i, entry.name)}
                  />
                ))}
              </Pie>
              <RechartsTooltip formatter={(value: any) => formatearEtiqueta(value, nombreY)} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Legend verticalAlign="bottom" height={36} />
            </PieChart>
          </ResponsiveContainer>
        );

      case 'rosa_polar': {
        const maxValRosa = Math.max(...datos.map((d: any) => Number(d[nombreY]) || 0), 1);
        const datosIguales = datos.map((d: any) => ({ ...d, _anguloIgual: 1, _valorReal: Number(d[nombreY]) || 0 }));
        return (
          <ResponsiveContainer width="100%" height={380}>
            <PieChart>
              <Pie
                data={datosIguales}
                cx="50%"
                cy="50%"
                innerRadius={0}
                outerRadius={130}
                dataKey="_anguloIgual"
                nameKey="name"
                isAnimationActive={false}
                label={bloque.mostrarValores ? ({ payload }: any) => `${payload.name}: ${formatearEtiqueta(payload._valorReal, nombreY)}` : false}
                shape={(props: any) => {
                  const { cx, cy, startAngle, endAngle, fill, payload, index } = props;
                  const val = Number(payload._valorReal) || 0;
                  const r = 22 + (val / maxValRosa) * 108;
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
                      style={{ cursor: 'pointer' }}
                      onClick={() => abrirSelectorDirecto(index, payload.name)}
                    />
                  );
                }}
              >
                {datosIguales.map((_, i) => (
                  <Cell key={`cell-rosa-${i}`} fill={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} />
                ))}
              </Pie>
              <RechartsTooltip formatter={(_: any, __: any, item: any) => [formatearEtiqueta(item?.payload?._valorReal, nombreY), item?.payload?.name]} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Legend verticalAlign="bottom" height={36} />
            </PieChart>
          </ResponsiveContainer>
        );
      }

      case 'torta_3d_rellena':
      case 'torta_4d': {
        const sinHueco = bloque.tipoGrafica === 'torta_3d_rellena';
        return (
          <div className="relative w-full flex flex-col items-center justify-center" style={{ height: 380 }}>
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <defs>
                  <filter id={`pie4d-admin-${bloque.id}`} x="-30%" y="-30%" width="160%" height="160%">
                    <feDropShadow dx="0" dy="14" stdDeviation="8" floodColor="#0f172a" floodOpacity="0.35" />
                  </filter>
                </defs>
                {/* Capa de profundidad 3D: legendType="none" y tooltipType="none" evitan duplicados en la leyenda */}
                <Pie 
                  data={datos} 
                  cx="50%" 
                  cy="55%" 
                  innerRadius={sinHueco ? 0 : 55} 
                  outerRadius={125} 
                  paddingAngle={sinHueco ? 0 : 4} 
                  dataKey={nombreY} 
                  isAnimationActive={false}
                  legendType="none"
                  tooltipType="none"
                >
                  {datos.map((_, i) => (
                    <Cell key={`depth-${i}`} fill={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} style={{ filter: 'brightness(0.62)' }} stroke="none" />
                  ))}
                </Pie>
                <Pie 
                  data={datos} 
                  cx="50%" 
                  cy="47%" 
                  innerRadius={sinHueco ? 0 : 55} 
                  outerRadius={125} 
                  paddingAngle={sinHueco ? 0 : 4} 
                  dataKey={nombreY} 
                  nameKey="name" 
                  stroke="#ffffff" 
                  strokeWidth={2}
                  style={{ filter: `url(#pie4d-admin-${bloque.id})` }}
                  label={bloque.mostrarValores ? ({ name, value }) => `${name} (${formatearEtiqueta(value, nombreY)})` : false}
                  isAnimationActive={false}
                >
                  {datos.map((entry: any, i: number) => (
                    <Cell 
                      key={`cell-4d-${i}`} 
                      fill={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} 
                      style={{ cursor: 'pointer' }}
                      onClick={() => abrirSelectorDirecto(i, entry.name)}
                    />
                  ))}
                </Pie>
                <RechartsTooltip formatter={(value: any) => formatearEtiqueta(value, nombreY)} contentStyle={{ borderRadius: '16px', border: '1px solid #e2e8f0', boxShadow: '0 20px 25px -5px rgb(0 0 0 / 0.2)' }} />
                <Legend verticalAlign="bottom" height={36} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        );
      }

      case 'barras_radiales':
        return (
          <ResponsiveContainer width="100%" height={360}>
            <RadialBarChart cx="50%" cy="50%" innerRadius="20%" outerRadius="95%" barSize={22} data={datos}>
              <RadialBar 
                background 
                dataKey={nombreY} 
                cornerRadius={10}
                isAnimationActive={false}
                onClick={(data: any, idx: number) => abrirSelectorDirecto(idx, data?.name || `Ítem ${idx + 1}`)}
              >
                {bloque.mostrarValores && <LabelList position="insideStart" fill="#1e293b" fontSize={11} fontWeight="bold" formatter={(v: any) => formatearEtiqueta(v, nombreY)} />}
              </RadialBar>
              <RechartsTooltip formatter={(value: any) => formatearEtiqueta(value, nombreY)} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Legend iconSize={10} layout="vertical" verticalAlign="middle" wrapperStyle={{ right: 0 }} />
            </RadialBarChart>
          </ResponsiveContainer>
        );

      case 'embudo':
        return (
          <ResponsiveContainer width="100%" height={360}>
            <FunnelChart>
              <RechartsTooltip formatter={(value: any) => formatearEtiqueta(value, nombreY)} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Funnel 
                dataKey={nombreY} 
                data={datos} 
                isAnimationActive={false}
                onClick={(data: any, idx: number) => abrirSelectorDirecto(idx, data?.name || `Ítem ${idx + 1}`)}
              >
                <LabelList position="right" fill="#292524" stroke="none" dataKey="name" style={{ fontSize: 12, fontWeight: 'bold' }} />
                {bloque.mostrarValores && <LabelList position="center" fill="#1e293b" stroke="none" dataKey={nombreY} formatter={(v: any) => formatearEtiqueta(v, nombreY)} style={{ fontSize: 12, fontWeight: 900 }} />}
              </Funnel>
            </FunnelChart>
          </ResponsiveContainer>
        );

      case 'treemap':
        return (
          <ResponsiveContainer width="100%" height={360}>
            <Treemap
              data={datos}
              dataKey={nombreY}
              nameKey="name"
              aspectRatio={4 / 3}
              stroke="#ffffff"
              isAnimationActive={false}
              fill={edicionActiva?.color_principal || '#93c5fd'}
              onClick={(node: any) => {
                const idx = typeof node?._idx === 'number' ? node._idx : (typeof node?.index === 'number' ? node.index : 0);
                abrirSelectorDirecto(idx, node?.name || `Ítem ${idx + 1}`);
              }}
            >
              <RechartsTooltip formatter={(value: any) => formatearEtiqueta(value, nombreY)} />
            </Treemap>
          </ResponsiveContainer>
        );

      case 'dispersion':
        return (
          <ResponsiveContainer width="100%" height={360}>
            <ScatterChart margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#e7e5e4" />
              <XAxis type="category" dataKey="name" name="Categoría" tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis type="number" dataKey={nombreY} name={nombreY} domain={[0, 'auto']} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip cursor={{ strokeDasharray: '3 3' }} formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), name]} />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              {nombresEjesY.map((nombre, i) => (
                <Scatter 
                  key={nombre} 
                  name={nombre} 
                  data={datos} 
                  isAnimationActive={false}
                  fill={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]}
                  onClick={() => abrirSelectorDirecto(i, nombre)}
                >
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                </Scatter>
              ))}
            </ScatterChart>
          </ResponsiveContainer>
        );

      case 'lineas':
      case 'lineas_curvas':
        return (
          <ResponsiveContainer width="100%" height={360}>
            <LineChart data={datos} margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis domain={[0, 'auto']} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), name]} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              {nombresEjesY.map((nombre, i) => (
                <Line 
                  type={bloque.tipoGrafica === 'lineas_curvas' ? 'natural' : 'linear'} 
                  key={nombre} 
                  dataKey={nombre} 
                  stroke={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} 
                  strokeWidth={3.5} 
                  dot={{ r: 4, strokeWidth: 2 }} 
                  activeDot={{ r: 6, onClick: () => abrirSelectorDirecto(i, nombre) }} 
                  onClick={() => abrirSelectorDirecto(i, nombre)}
                  isAnimationActive={false}
                >
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                </Line>
              ))}
            </LineChart>
          </ResponsiveContainer>
        );

      case 'area':
      case 'area_apilada':
        return (
          <ResponsiveContainer width="100%" height={360}>
            <AreaChart data={datos} margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis domain={[0, 'auto']} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), name]} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              {nombresEjesY.map((nombre, i) => (
                <Area 
                  type="monotone" 
                  key={nombre} 
                  dataKey={nombre} 
                  stackId={bloque.tipoGrafica === 'area_apilada' ? '1' : undefined} 
                  stroke={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} 
                  fill={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} 
                  fillOpacity={0.65} 
                  onClick={() => abrirSelectorDirecto(i, nombre)}
                  isAnimationActive={false}
                >
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                </Area>
              ))}
            </AreaChart>
          </ResponsiveContainer>
        );

      case 'combinada':
        return (
          <ResponsiveContainer width="100%" height={360}>
            <ComposedChart data={datos} margin={{ top: 28, right: 30, left: 0, bottom: 5 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e7e5e4" />
              <XAxis dataKey="name" tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} />
              <YAxis domain={[0, 'auto']} tickLine={false} axisLine={false} tick={{ fill: '#78716c', fontSize: 12 }} tickFormatter={(v) => sufijoPorcentaje ? `${v}%` : v} />
              <RechartsTooltip formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), name]} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              {nombresEjesY.map((nombre, i) => {
                const color = bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length];
                if (i === nombresEjesY.length - 1 && nombresEjesY.length > 1) {
                  return (
                    <Line key={nombre} type="monotone" dataKey={nombre} stroke={color} strokeWidth={3} dot={{ r: 4 }} isAnimationActive={false} onClick={() => abrirSelectorDirecto(i, nombre)}>
                      {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#292524', fontSize: 11, fontWeight: 'bold' }} />}
                    </Line>
                  );
                }
                return (
                  <Bar key={nombre} dataKey={nombre} fill={color} radius={[6, 6, 0, 0]} isAnimationActive={false} onClick={() => abrirSelectorDirecto(i, nombre)}>
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
          <ResponsiveContainer width="100%" height={360}>
            <RadarChart cx="50%" cy="50%" outerRadius="78%" data={datos}>
              <PolarGrid gridType={bloque.tipoGrafica === 'radar_relleno' ? 'circle' : 'polygon'} stroke="#cbd5e1" />
              <PolarAngleAxis dataKey="name" tick={{ fill: '#475569', fontSize: 12, fontWeight: 'bold' }} />
              <PolarRadiusAxis angle={30} domain={[0, 'auto']} tick={{ fill: '#94a3b8', fontSize: 10 }} />
              <RechartsTooltip formatter={(value: any, name: any) => [formatearEtiqueta(value, String(name)), name]} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
              <Legend wrapperStyle={{ paddingTop: '20px' }} />
              {nombresEjesY.map((nombre, i) => (
                <Radar 
                  key={nombre} 
                  name={nombre} 
                  dataKey={nombre} 
                  stroke={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} 
                  strokeWidth={2.5} 
                  fill={bloque.coloresCustom?.[i] || COLORES_GRAFICAS_DEFAULT[i % COLORES_GRAFICAS_DEFAULT.length]} 
                  fillOpacity={bloque.tipoGrafica === 'radar_relleno' ? 0.75 : 0.45} 
                  onClick={() => abrirSelectorDirecto(i, nombre)}
                  isAnimationActive={false}
                >
                  {bloque.mostrarValores && <LabelList dataKey={nombre} position="top" formatter={(v: any) => formatearEtiqueta(v, nombre)} style={{ fill: '#1e293b', fontSize: 11, fontWeight: 'bold' }} />}
                </Radar>
              ))}
            </RadarChart>
          </ResponsiveContainer>
        );

      case 'mapa_colombia':
        const valorMaximo = Math.max(...datos.map((d: any) => Number(d.valorMapa) || 0), 1);
        const colorMapa = bloque.coloresCustom?.[0] || edicionActiva?.color_principal || "#93c5fd";
        const escalaColor = scaleLinear<string>().domain([0, valorMaximo]).range(["#f1f5f9", colorMapa]); 
        
        const estiloMapa: any = {
          default: { outline: "none" },
          hover: { fill: "#fde047", stroke: "#0f172a", strokeWidth: 1.2, outline: "none", cursor: "pointer", transition: "all 150ms" },
          pressed: { fill: "#facc15", outline: "none" }
        };

        return (
          <div className="w-full relative flex flex-col md:flex-row items-center rounded-2xl border border-stone-200 overflow-hidden bg-transparent" style={{ minHeight: '400px' }}>
            {tooltipMapa && (
              <div 
                className="fixed z-9999 pointer-events-none bg-slate-900/95 text-white px-3.5 py-2 rounded-xl shadow-xl border border-slate-700 text-xs font-sans transform -translate-x-1/2 -translate-y-full -mt-3 transition-opacity duration-75"
                style={{ left: tooltipMapa.x, top: tooltipMapa.y }}
              >
                <span className="block text-[10px] uppercase tracking-widest text-slate-400 font-bold">{tooltipMapa.nombre}</span>
                <span className="block text-sm font-black text-amber-300">{nombreY || 'Valor'}: {tooltipMapa.valor}</span>
              </div>
            )}
            <div className="w-full" style={{ height: '400px' }}>
              <ComposableMap projection="geoMercator" projectionConfig={{ scale: 1800, center: [-74, 4.5] }} className="w-full h-full bg-transparent">
                <ZoomableGroup center={[-74, 4.5]} zoom={1} maxZoom={5}>
                  <Geographies geography="/colombia.json">
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
                            onClick={() => abrirSelectorDirecto(0, 'Color Principal del Mapa')}
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
              <div className="w-full md:w-64 max-h-95 overflow-y-auto bg-white/90 backdrop-blur border-t md:border-t-0 md:border-l border-stone-200 p-4 font-sans text-xs space-y-2 shrink-0">
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
  // RENDERIZADOR INFALIBLE DE BOTONES INLINE DENTRO DEL TEXTO (SOPORTA SOLIDO, PILDORA, CONTORNO, 3D Y 4D)
  const renderizarBotonesEnTexto = (html: string) => {
    if (!html) return '';
    const htmlSinNbsp = html.replace(/&nbsp;|&#160;|\u00A0/gi, ' ');
    const htmlConSaltosPreservados = htmlSinNbsp
      .replace(/<p(\s+[^>]*)?>\s*<br\s*\/?>\s*<\/p>/gi, '<p$1>&nbsp;</p>')
      .replace(/<p(\s+[^>]*)?>\s*<\/p>/gi, '<p$1>&nbsp;</p>');

    return htmlConSaltosPreservados.replace(/\[\[BOTON\|([\s\S]*?)\]\]/g, (_, contenidoInterno) => {
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
      }

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
      const targetAttr = esLinkInterno ? '' : 'target="_blank" rel="noreferrer"';

      return `<a href="${url}" ${targetAttr} data-inline-btn="true" style="${extraStyle} padding: 9px 22px !important; font-family: sans-serif !important; font-weight: 900 !important; font-size: 12px !important; line-height: 1.2 !important; text-transform: uppercase !important; letter-spacing: 0.06em !important; text-decoration: none !important; display: inline-block !important; margin: 6px 6px !important; vertical-align: middle !important; cursor: pointer !important; transition: transform 0.2s ease, filter 0.2s ease !important;">${texto}</a>`;
    });
  };

  const textoVacioQuill = (html?: string) => {
    if (!html) return true;
    const limpio = html.replace(/<[^>]*>/g, '').trim();
    return limpio === '' && !html.includes('<img');
  };

  const contenidoPortadaPorDefecto = `
    <h1 class="text-6xl font-black text-center">Programas de Ingeniería en Colombia</h1>
    <p class="text-2xl text-center italic">Análisis oficial y directorio nacional de la educación superior en ingeniería.</p>
  `;

  return (
    <div className={`${generandoPdfAdmin ? 'min-h-screen w-full overflow-visible block bg-[#fdfbf7]' : 'flex h-screen bg-slate-50 font-sans text-slate-900 antialiased overflow-hidden relative'}`}>
      <style dangerouslySetInnerHTML={{ __html: ESTILOS_FUENTES_QUILL }} />
      
      {mensaje && (
        <div className={`fixed top-4 right-4 z-50 px-6 py-3 rounded-xl font-bold shadow-lg border animate-in slide-in-from-top-4 no-print ${mensaje.tipo === 'exito' ? 'bg-green-50 text-green-700 border-green-200' : 'bg-red-50 text-red-700 border-red-200'}`}>
          {mensaje.texto}
        </div>
      )}

      {/* MENÚ LATERAL COLAPSABLE (BIBLIOTECA DE EDICIONES) */}
      {menuLateralAbierto && !generandoPdfAdmin && (
        <aside className="w-80 bg-white border-r border-slate-200 flex flex-col h-full z-10 shrink-0 shadow-sm transition-all duration-300 no-print">
          <div className="px-6 py-8 border-b border-slate-100 flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <div className="p-3 bg-blue-600 text-white rounded-xl shadow-lg shadow-blue-600/30">
                <Book size={26} strokeWidth={2.5} />
              </div>
              <div>
                <h1 className="text-2xl font-extrabold tracking-tight text-slate-900">ACOFI</h1>
                <p className="text-xs font-bold text-blue-600 uppercase tracking-widest mt-0.5">Gestor CMS</p>
              </div>
            </div>
            <div className="flex items-center gap-1">
              <button 
                onClick={() => setMenuLateralAbierto(false)} 
                className="p-2 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-colors cursor-pointer" 
                title="Ocultar menú lateral para ampliar el editor"
              >
                <PanelLeftClose size={20} strokeWidth={2.2} />
              </button>
              <button 
                onClick={cerrarSesion} 
                className="p-2 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded-lg transition-colors cursor-pointer" 
                title="Cerrar Sesión"
              >
                <Power size={20} strokeWidth={2.2} />
              </button>
            </div>
          </div>
          <div className="flex-1 overflow-y-auto p-5 space-y-3">
            <h2 className="text-[12px] font-black text-slate-400 uppercase tracking-widest mb-4 px-1">Biblioteca</h2>
            {ediciones.map((ed) => {
              const isActiva = edicionActiva?.id === ed.id;
              return (
                <button key={ed.id} onClick={() => cambiarEdicionActiva(ed)} className={`w-full text-left px-5 py-4 rounded-2xl transition-all border-2 cursor-pointer ${isActiva ? 'bg-slate-900 border-slate-900 text-white shadow-xl' : 'bg-white border-slate-100 text-slate-600 hover:border-slate-300'}`}>
                  <div className={`font-black text-xl mb-2 flex items-center justify-between ${isActiva ? 'text-white' : 'text-slate-800'}`}>
                    <span>{ed.anio}</span>
                    <span className={`text-[10px] px-2 py-0.5 rounded-full uppercase tracking-wider font-bold ${ed.estado === 'publicado' ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-amber-500/20 text-amber-300 border border-amber-500/30'}`}>
                      {ed.estado || 'borrador'}
                    </span>
                  </div>
                  <div className={`text-xs font-medium w-40 truncate ${isActiva ? 'text-slate-400' : 'text-slate-500'}`}>{ed.titulo}</div>
                </button>
              );
            })}
          </div>
          <div className="p-5 border-t border-slate-100 bg-slate-50 space-y-3">
            <button onClick={descargarRespaldoJson} disabled={!edicionActiva} className="w-full flex justify-center items-center space-x-2 bg-emerald-50 border-2 border-emerald-200 text-emerald-800 py-3 px-4 rounded-xl text-xs font-bold hover:bg-emerald-100 transition-all shadow-xs cursor-pointer">
              <FileJson size={16} /> <span>Descargar Respaldo JSON</span>
            </button>
            <button onClick={intentarDuplicarEdicion} disabled={guardando || !edicionActiva} className="w-full flex justify-center items-center space-x-2 bg-white border-2 border-slate-200 text-slate-700 py-3.5 px-4 rounded-xl text-sm font-bold hover:bg-blue-50 hover:border-blue-200 hover:text-blue-700 transition-all shadow-sm cursor-pointer">
              <Copy size={18} /> <span>Clonar Edición</span>
            </button>
          </div>
        </aside>
      )}

      {/* VISTA DE IMPRESIÓN / EXPORTACIÓN PDF OFICIAL DESDE EL ADMIN (FLUJO NATIVO PARA QUE IMPRIMA LAS 120 PÁGINAS IDÉNTICAS AL LIBRO) */}
      {generandoPdfAdmin && (
        <div className="w-full bg-[#fdfbf7] font-serif text-stone-900 print-book-wrapper">
          <div className="max-w-5xl mx-auto">
            
            {/* Portada en PDF */}
            <div className="print-page-sheet flex flex-col items-center justify-center text-center px-10" style={{ backgroundColor: configuracion.colorFondoPortada || '#fdfbf7' }}>
              {configuracion.logoUrl && (
                <img src={configuracion.logoUrl} alt="Logo" style={{ height: `${configuracion.logoSize ?? 56}px` }} className="w-auto object-contain mb-8" />
              )}
              <div className="w-16 h-1 mb-10" style={{ backgroundColor: edicionActiva?.color_principal || '#1d4ed8' }}></div>
              <div 
                className="w-full prose prose-stone max-w-none [&_h1]:text-5xl [&_h1]:font-normal [&_h1]:mb-6 [&_p]:text-xl [&_p]:italic [&_p]:text-stone-600"
                dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(configuracion.portadaContenido || contenidoPortadaPorDefecto) }}
              />
              <div className="mt-20 flex flex-col items-center">
                <span className="text-xs font-sans font-bold tracking-[0.3em] uppercase text-stone-400 mb-2">ACOFI / Edición Oficial {edicionActiva?.anio}</span>
              </div>
            </div>

            {/* Todas las Páginas del Libro en PDF */}
            {paginas.map((p, pIdx) => {
              const tieneT = p.titulo.trim() !== '';
              const tieneSub = (p.subtitulo || '').trim() !== '';
              const mostrarNum = p.mostrarNumero !== false && tieneT;
              const cabeceraMostrar = !textoVacioQuill(p.cabeceraHtml)
                ? p.cabeceraHtml
                : (configuracion.mostrarCabeceraGlobal && !textoVacioQuill(configuracion.cabeceraGlobalHtml) ? configuracion.cabeceraGlobalHtml : '');
              const pieMostrar = !textoVacioQuill(p.piePaginaHtml)
                ? p.piePaginaHtml
                : (configuracion.mostrarPiePaginaGlobal && !textoVacioQuill(configuracion.piePaginaGlobalHtml) ? configuracion.piePaginaGlobalHtml : '');

              return (
                <div key={p.id} className="print-page-sheet flex flex-col" style={{ backgroundColor: p.colorFondo || '#fdfbf7' }}>
                  {cabeceraMostrar && (
                    <div className="mb-6 pb-3 border-b border-stone-200 text-xs opacity-75" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(cabeceraMostrar) }} />
                  )}
                  {(tieneT || tieneSub) && (
                    <div className="mb-8 border-b border-stone-200 pb-4 print-avoid-break">
                      <div className="flex items-center gap-4">
                        {mostrarNum && <span className="text-4xl font-black opacity-30">{pIdx + 1}.</span>}
                        {tieneT && <h2 className="text-3xl font-black tracking-tight">{p.titulo}</h2>}
                      </div>
                      {tieneSub && <p className="text-base text-stone-500 font-serif italic mt-2">{p.subtitulo}</p>}
                    </div>
                  )}
                  <div className="space-y-8 flex-1">
                    {p.bloques.map((b) => (
                      <div key={b.id} className="w-full print-avoid-break">
                        {b.tipo === 'texto' && (
                          b.columnas === 2 || b.columnas === 3 ? (
                            <div 
                              className={`grid ${b.columnas === 2 ? 'grid-cols-2 gap-8' : 'grid-cols-3 gap-6'} p-4 rounded-2xl items-start`}
                              style={{ backgroundColor: b.colorFondo === 'transparent' ? 'transparent' : (b.colorFondo || 'transparent') }}
                            >
                              <div className="prose prose-stone max-w-none text-stone-700 leading-normal" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.contenido) }} />
                              <div className="prose prose-stone max-w-none text-stone-700 leading-normal border-l border-stone-200 pl-6" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.contenidoCol2 || '') }} />
                              {b.columnas === 3 && <div className="prose prose-stone max-w-none text-stone-700 leading-normal border-l border-stone-200 pl-6" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.contenidoCol3 || '') }} />}
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
                            className={`py-4 px-4 rounded-2xl flex flex-wrap gap-10 items-end ${
                              b.alineacion === 'centro' ? 'justify-center text-center' :
                              b.alineacion === 'derecha' ? 'justify-end text-right' :
                              b.alineacion === 'espaciado' ? 'justify-around text-left' :
                              'justify-start text-left'
                            }`}
                            style={{ 
                              backgroundColor: b.colorFondo === 'transparent' ? 'transparent' : (b.colorFondo || 'transparent'),
                              paddingTop: `${(b.espacioGeneralSuperior ?? 0) + 16}px`,
                              paddingBottom: `${(b.espacioGeneralInferior ?? 0) + 16}px`
                            }}
                          >
                            {(b.firmas || []).map((f, i) => (
                              <div 
                                key={i} 
                                className="flex flex-col space-y-1 min-w-52"
                                style={{ marginTop: `${f.espacioSuperior || 0}px` }}
                              >
                                {f.url && (
                                  <img 
                                    src={f.url} 
                                    alt={f.nombre} 
                                    style={{ height: `${f.tamanoFirma ?? 72}px`, maxHeight: `${f.tamanoFirma ?? 72}px`, width: 'auto' }} 
                                    className={`object-contain mb-2 ${b.alineacion === 'centro' ? 'mx-auto' : b.alineacion === 'derecha' ? 'ml-auto' : 'mr-auto'}`} 
                                  />
                                )}
                                <span className="font-sans font-black text-stone-900 text-base leading-snug">{f.nombre}</span>
                                {f.cargo && <span className="font-sans text-stone-600 text-sm leading-snug">{f.cargo}</span>}
                                {f.infoAdicional && <span className="font-sans text-stone-500 text-xs italic leading-snug">{f.infoAdicional}</span>}
                              </div>
                            ))}
                          </div>
                        )}
                        {b.tipo === 'boton' && (
                          <div className={`py-3 flex ${b.alineacion === 'centro' ? 'justify-center' : b.alineacion === 'derecha' ? 'justify-end' : 'justify-start'}`}>
                            <span
                              style={{
                                backgroundColor: b.estilo === 'contorno' ? 'transparent' : b.colorBoton,
                                color: b.estilo === 'contorno' ? b.colorBoton : b.colorTexto,
                                borderColor: b.colorBoton,
                                boxShadow: b.estilo === '4d' ? `0 0 20px 3px ${b.colorIluminacion || b.colorBoton}` : undefined
                              }}
                              className={`px-6 py-2.5 font-sans font-black text-xs uppercase tracking-wider inline-block ${b.estilo === 'pildora' || b.estilo === '4d' ? 'rounded-full' : 'rounded-xl'}`}
                            >
                              {b.texto}
                            </span>
                          </div>
                        )}
                        {b.tipo === 'tabla' && !b.oculto && (
                          <div className="border border-stone-300 rounded-2xl overflow-hidden my-4" style={{ backgroundColor: b.colorFondo || '#ffffff' }}>
                            <div className="bg-stone-100 p-3 font-sans font-bold text-xs uppercase">{b.nombre}</div>
                            <table className="w-full text-left font-sans border-collapse" style={{ fontSize: `${b.tamanoTextoTabla ?? 13}px` }}>
                              <thead>
                                <tr className="bg-stone-200/70 border-b border-stone-300" style={{ backgroundColor: b.coloresFilas?.[0] || '#f5f5f4' }}>
                                  {b.data[0]?.map((c, ci) => <th key={ci} className="p-2 font-bold uppercase">{c}</th>)}
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-stone-200">
                                {b.data.slice(1).map((f, fi) => (
                                  <tr key={fi} style={{ backgroundColor: b.coloresFilas?.[fi + 1] || 'transparent' }}>
                                    {f.map((cell, cii) => <td key={cii} className={`p-2 ${cii === 0 ? 'font-semibold' : ''}`}>{cell}</td>)}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                            {!textoVacioQuill(b.pieBloque) && <div className="p-2.5 bg-stone-50 text-xs border-t border-stone-200" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.pieBloque || '') }} />}
                          </div>
                        )}
                        {b.tipo === 'grafica' && (
                          <div className="border border-stone-200 rounded-2xl p-6 my-4" style={{ backgroundColor: b.colorFondo || '#ffffff' }}>
                            <h3 className="font-sans font-bold uppercase tracking-widest text-xs text-center mb-4">{b.nombre}</h3>
                            {renderizarGrafica(b)}
                            {!textoVacioQuill(b.pieBloque) && <div className="mt-3 pt-2 border-t border-stone-200 text-xs" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(b.pieBloque || '') }} />}
                          </div>
                        )}
                        {b.tipo === 'imagen' && b.url && (
                          <div className="my-4 text-center">
                            <img src={b.url} alt="Imagen" className="max-h-80 mx-auto rounded-xl object-contain" />
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                  {pieMostrar && (
                    <div className="mt-8 pt-3 border-t border-stone-200 text-xs opacity-75" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(pieMostrar) }} />
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}

      {!generandoPdfAdmin && (
        <main className="flex-1 overflow-y-auto bg-slate-50 relative no-print">
          {edicionActiva ? (
            <div className={`${menuLateralAbierto ? 'max-w-6xl' : 'max-w-full'} mx-auto p-6 lg:p-10 pb-32 transition-all duration-300`}>
              
              <div className="flex flex-col md:flex-row md:items-end justify-between mb-10 gap-6 border-b border-slate-200 pb-6">
                <div className="flex items-start gap-4">
                  {!menuLateralAbierto && (
                    <button
                      onClick={() => setMenuLateralAbierto(true)}
                      className="mt-1 p-3 bg-white border-2 border-slate-200 text-slate-700 hover:border-blue-500 hover:text-blue-600 rounded-xl shadow-xs transition-all flex items-center gap-2 text-xs font-bold shrink-0 cursor-pointer"
                      title="Mostrar Biblioteca de Ediciones"
                    >
                      <PanelLeftOpen size={18} />
                      <span>Biblioteca</span>
                    </button>
                  )}
                  <div>
                    <h2 className="text-4xl font-black text-slate-900 tracking-tight mb-2 flex items-center gap-3">
                      <Edit3 size={32} className="text-blue-600" /> Edición {edicionActiva.anio}
                    </h2>
                    <div className="flex items-center space-x-3 text-sm font-bold text-slate-400 uppercase tracking-widest">
                      <Database size={16} /><span>Motor de Bloques Relacional</span>
                      <span className="text-slate-300">•</span>
                      <select
                        value={edicionActiva.estado || 'borrador'}
                        onChange={(e) => actualizarEstadoEdicion(e.target.value as 'publicado' | 'borrador')}
                        className={`text-xs font-black uppercase px-3 py-1 rounded-full border outline-none cursor-pointer ${edicionActiva.estado === 'publicado' ? 'bg-emerald-50 text-emerald-700 border-emerald-300' : 'bg-amber-50 text-amber-700 border-amber-300'}`}
                      >
                        <option value="borrador">⚠️ Borrador</option>
                        <option value="publicado">✅ Publicado</option>
                      </select>
                    </div>
                  </div>
                </div>
                <div className="flex items-center gap-2.5 flex-wrap">
                  <button
                    type="button"
                    onClick={descargarRespaldoJson}
                    className="flex items-center gap-2 text-sm font-bold text-slate-700 bg-white border border-slate-200 px-4 py-3 rounded-xl hover:border-slate-400 hover:bg-slate-50 transition-all shadow-xs cursor-pointer"
                    title="Descargar copia de seguridad completa en formato JSON"
                  >
                    <FileJson size={18} className="text-emerald-600" /> <span>Descargar JSON</span>
                  </button>
                  <button
                    type="button"
                    onClick={descargarLibroPdfAdmin}
                    disabled={generandoPdfAdmin}
                    className="flex items-center gap-2 text-sm font-bold text-emerald-800 bg-emerald-50 border border-emerald-200 px-5 py-3 rounded-xl hover:bg-emerald-100 transition-all shadow-xs cursor-pointer"
                    title="Exportar todo el libro completo en PDF editorial listo para imprimir"
                  >
                    <Printer size={18} /> <span>{generandoPdfAdmin ? 'Preparando PDF...' : 'Descargar PDF'}</span>
                  </button>
                  <a href={`/?edicion=${edicionActiva.anio}`} target="_blank" rel="noreferrer" className="flex items-center gap-2 text-sm font-bold text-blue-700 bg-blue-100/50 border border-blue-200 px-6 py-3 rounded-xl hover:bg-blue-100 hover:shadow-md transition-all">
                    <Eye size={18} /> <span>Ver Libro Público</span>
                  </a>
                </div>
              </div>

              <div className="flex space-x-8 mb-10 border-b border-slate-200">
                <button onClick={() => setPestañaActiva('diseno')} className={`pb-4 text-sm font-bold transition-colors relative flex items-center gap-2 cursor-pointer ${pestañaActiva === 'diseno' ? 'text-blue-600' : 'text-slate-500 hover:text-slate-900'}`}><Layout size={18} /> Propiedades de Revista {pestañaActiva === 'diseno' && <div className="absolute bottom-0 left-0 w-full h-1 bg-blue-600 rounded-t-full"></div>}</button>
                <button onClick={() => setPestañaActiva('contenido')} className={`pb-4 text-sm font-bold transition-colors relative flex items-center gap-2 cursor-pointer ${pestañaActiva === 'contenido' ? 'text-blue-600' : 'text-slate-500 hover:text-slate-900'}`}><Type size={18} /> Constructor de Páginas {pestañaActiva === 'contenido' && <div className="absolute bottom-0 left-0 w-full h-1 bg-blue-600 rounded-t-full"></div>}</button>
              </div>

              {pestañaActiva === 'diseno' && (
                <form onSubmit={manejarGuardadoManual} className="space-y-8 animate-in fade-in duration-500">
                  
                  <div className="bg-white rounded-3xl shadow-sm border border-slate-200">
                    <div className="bg-slate-50 border-b border-slate-200 p-6 flex justify-between items-center rounded-t-3xl">
                      <h3 className="text-lg font-black text-slate-800">1. Identidad Principal y Colores</h3>
                      <button 
                        type="submit" 
                        disabled={guardando} 
                        className="bg-blue-600 text-white px-5 py-2.5 rounded-xl font-bold text-sm shadow hover:bg-blue-700 transition-all flex items-center gap-2 cursor-pointer"
                      >
                        <Save size={16} /> <span>Guardar Cambios</span>
                      </button>
                    </div>
                    <div className="p-8 space-y-8">
                      
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                        <div className="md:col-span-2">
                          <label className="block text-sm font-bold text-slate-700 mb-2 uppercase tracking-wide">Nombre de la Edición (Interno)</label>
                          <input 
                            type="text" 
                            value={edicionActiva.titulo || ''} 
                            onChange={(e) => {
                              const nuevo = { ...edicionActiva, titulo: e.target.value };
                              setEdicionActiva(nuevo);
                              supabase.from('ediciones').update({ titulo: e.target.value }).eq('id', edicionActiva.id);
                            }} 
                            className="w-full px-5 py-4 bg-slate-50 border-2 border-slate-200 rounded-xl font-semibold text-lg outline-none focus:border-blue-500" 
                            required 
                          />
                        </div>

                        <div>
                          <label className="block text-sm font-bold text-slate-700 mb-2 uppercase tracking-wide">Estado de Publicación</label>
                          <select
                            value={edicionActiva.estado || 'borrador'}
                            onChange={(e) => actualizarEstadoEdicion(e.target.value as 'publicado' | 'borrador')}
                            className="w-full px-5 py-4 bg-slate-50 border-2 border-slate-200 rounded-xl font-bold text-base outline-none focus:border-blue-500 cursor-pointer"
                          >
                            <option value="borrador">Borrador (Privado)</option>
                            <option value="publicado">Publicado (Público)</option>
                          </select>
                        </div>
                      </div>
                      
                      <div>
                        <label className="block text-sm font-bold text-slate-700 mb-2 uppercase tracking-wide">Textos de la Portada (Editor Completo)</label>
                        <div className="bg-white border-2 border-slate-200 rounded-xl [&_.ql-toolbar]:border-none [&_.ql-toolbar]:bg-slate-50 [&_.ql-toolbar]:border-b [&_.ql-toolbar]:border-slate-200 [&_.ql-toolbar]:rounded-t-xl [&_.ql-container]:border-none [&_.ql-editor]:min-h-37.5">
                          <ReactQuill 
                            theme="snow" 
                            modules={quillModules} 
                            value={configuracion.portadaContenido !== undefined ? configuracion.portadaContenido : contenidoPortadaPorDefecto} 
                            onChange={(val) => actualizarConfiguracion('portadaContenido', val)} 
                          />
                        </div>
                        <p className="text-xs text-slate-500 mt-2">Personaliza el título, subtítulo, tamaños y colores que aparecerán al inicio de la revista.</p>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-3 gap-8 pt-4 border-t border-slate-100">
                        <div>
                          <label className="block text-sm font-bold text-slate-700 mb-2 uppercase tracking-wide">Año Comercial</label>
                          <input 
                            type="number" 
                            value={edicionActiva.anio || ''} 
                            onChange={(e) => {
                               setEdicionActiva({ ...edicionActiva, anio: e.target.value });
                               supabase.from('ediciones').update({ anio: e.target.value }).eq('id', edicionActiva.id);
                            }} 
                            className="w-full px-5 py-4 bg-slate-50 border-2 border-slate-200 rounded-xl font-semibold text-lg outline-none focus:border-blue-500" 
                            required 
                          />
                        </div>

                        <div>
                          <label className="block text-sm font-bold text-slate-700 mb-2 uppercase tracking-wide">Color Tema (HEX)</label>
                          <div className="flex gap-3 items-center">
                            <input 
                              type="color" 
                              value={edicionActiva.color_principal || '#1d4ed8'} 
                              onChange={(e) => actualizarColorTema(e.target.value)} 
                              className="w-14 h-14 cursor-pointer rounded-xl border-2 shrink-0" 
                            />
                            <input 
                              type="text" 
                              value={edicionActiva.color_principal || ''} 
                              onChange={(e) => actualizarColorTema(e.target.value)} 
                              className="w-full px-4 py-4 bg-slate-50 border-2 border-slate-200 rounded-xl uppercase font-mono font-bold text-base outline-none focus:border-blue-500" 
                              required 
                            />
                          </div>
                        </div>

                        <div>
                          <label className="block text-sm font-bold text-slate-700 mb-2 uppercase tracking-wide">Fondo de Portada</label>
                          <div className="flex gap-3 items-center">
                            <input 
                              type="color" 
                              value={configuracion.colorFondoPortada || '#fdfbf7'} 
                              onChange={(e) => actualizarConfiguracion('colorFondoPortada', e.target.value)} 
                              className="w-14 h-14 cursor-pointer rounded-xl border-2 shrink-0" 
                            />
                            <input 
                              type="text" 
                              value={configuracion.colorFondoPortada || '#fdfbf7'} 
                              onChange={(e) => actualizarConfiguracion('colorFondoPortada', e.target.value)} 
                              className="w-full px-4 py-4 bg-slate-50 border-2 border-slate-200 rounded-xl uppercase font-mono font-bold text-base outline-none focus:border-blue-500" 
                            />
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="bg-white rounded-3xl shadow-sm border border-slate-200 overflow-hidden">
                    <div className="bg-slate-50 border-b border-slate-200 p-6 flex items-center gap-3">
                      <ImagePlus className="text-blue-600" />
                      <h3 className="text-lg font-black text-slate-800">2. Multimedia y Branding Global</h3>
                    </div>
                    <div className="p-8">
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                        <RenderUploadConfig campo="imagenFondoPortada" label="Imagen Fondo (Portada)" placeholder="URL de imagen de fondo..." />
                        
                        <div className="space-y-4">
                          <RenderUploadConfig campo="logoUrl" label="Logo Superior (Páginas)" placeholder="URL del logo en PNG/SVG..." />
                          {configuracion.logoUrl && (
                            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-4">
                              <div>
                                <label className="flex justify-between text-xs font-bold text-slate-600 uppercase mb-2">
                                  <span>Tamaño del Logo</span>
                                  <span className="text-blue-600">{configuracion.logoSize ?? 56}px</span>
                                </label>
                                <input type="range" min="20" max="250" value={configuracion.logoSize ?? 56} onChange={(e) => actualizarConfiguracion('logoSize', Number(e.target.value))} className="w-full accent-blue-600" />
                              </div>
                              <div>
                                <label className="flex justify-between text-xs font-bold text-slate-600 uppercase mb-2">
                                  <span>Luminosidad</span>
                                  <span className="text-blue-600">{configuracion.logoBrightness ?? 100}%</span>
                                </label>
                                <input type="range" min="0" max="200" value={configuracion.logoBrightness ?? 100} onChange={(e) => actualizarConfiguracion('logoBrightness', Number(e.target.value))} className="w-full accent-blue-600" />
                              </div>
                            </div>
                          )}
                        </div>

                        <div className="space-y-4">
                          <RenderUploadConfig campo="marcaAguaUrl" label="Marca de Agua Central" placeholder="URL de la marca de agua..." />
                          {configuracion.marcaAguaUrl && (
                            <div className="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-4">
                              <div>
                                <label className="flex justify-between text-xs font-bold text-slate-600 uppercase mb-2">
                                  <span>Tamaño (Ancho Pantalla)</span>
                                  <span className="text-blue-600">{configuracion.marcaAguaSize ?? 40}%</span>
                                </label>
                                <input type="range" min="10" max="100" value={configuracion.marcaAguaSize ?? 40} onChange={(e) => actualizarConfiguracion('marcaAguaSize', Number(e.target.value))} className="w-full accent-blue-600" />
                              </div>
                              <div>
                                <label className="flex justify-between text-xs font-bold text-slate-600 uppercase mb-2">
                                  <span>Visibilidad (Opacidad)</span>
                                  <span className="text-blue-600">{configuracion.marcaAguaOpacity ?? 5}%</span>
                                </label>
                                <input type="range" min="1" max="100" value={configuracion.marcaAguaOpacity ?? 5} onChange={(e) => actualizarConfiguracion('marcaAguaOpacity', Number(e.target.value))} className="w-full accent-blue-600" />
                              </div>
                            </div>
                          )}
                        </div>

                      </div>

                      {configuracion.marcaAguaUrl && (
                        <div className="mt-8 pt-8 border-t border-slate-200">
                          <h4 className="text-sm font-bold text-slate-800 uppercase tracking-widest mb-4">¿En dónde mostrar la Marca de Agua?</h4>
                          <div className="flex items-center gap-6 mb-6">
                            <label className="flex items-center gap-2 cursor-pointer font-bold text-slate-700">
                              <input 
                                type="radio" 
                                name="modoMarcaAgua" 
                                checked={configuracion.modoMarcaAgua === 'todas'} 
                                onChange={() => actualizarConfiguracion('modoMarcaAgua', 'todas')}
                                className="w-4 h-4 text-blue-600"
                              /> En todas las páginas
                            </label>
                            <label className="flex items-center gap-2 cursor-pointer font-bold text-slate-700">
                              <input 
                                type="radio" 
                                name="modoMarcaAgua" 
                                checked={configuracion.modoMarcaAgua === 'seleccionadas'} 
                                onChange={() => actualizarConfiguracion('modoMarcaAgua', 'seleccionadas')}
                                className="w-4 h-4 text-blue-600"
                              /> Elegir páginas específicas
                            </label>
                          </div>

                          {configuracion.modoMarcaAgua === 'seleccionadas' && (
                            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 bg-slate-50 p-6 rounded-xl border border-slate-200">
                              {paginas.map((p, idx) => (
                                <label key={p.id} className="flex items-center gap-2 text-sm font-medium text-slate-600 cursor-pointer bg-white p-3 rounded-lg border border-slate-200 hover:border-blue-400">
                                  <input 
                                    type="checkbox" 
                                    checked={(configuracion.paginasMarcaAgua || []).includes(p.id)}
                                    onChange={(e) => {
                                      let arr = [...(configuracion.paginasMarcaAgua || [])];
                                      if (e.target.checked) arr.push(p.id);
                                      else arr = arr.filter(id => id !== p.id);
                                      actualizarConfiguracion('paginasMarcaAgua', arr);
                                    }}
                                    className="w-4 h-4 rounded text-blue-600 focus:ring-blue-500"
                                  />
                                  {p.titulo || `Página ${idx + 1}`}
                                </label>
                              ))}
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  </div>

                  <div className="bg-white rounded-3xl shadow-sm border border-slate-200">
                    <div className="bg-slate-50 border-b border-slate-200 p-6 flex items-center gap-3 rounded-t-3xl">
                      <Type className="text-blue-600" />
                      <h3 className="text-lg font-black text-slate-800">3. Cabecera y Pie de Página Globales (Opcional)</h3>
                    </div>
                    <div className="p-8 space-y-8">
                      <div className="space-y-3">
                        <label className="flex items-center gap-3 text-sm font-bold text-slate-700 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={!!configuracion.mostrarCabeceraGlobal}
                            onChange={(e) => actualizarConfiguracion('mostrarCabeceraGlobal', e.target.checked)}
                            className="w-4 h-4 rounded text-blue-600"
                          />
                          Habilitar Cabecera Global en todas las páginas de contenido
                        </label>
                        {configuracion.mostrarCabeceraGlobal && (
                          <div className="bg-white border-2 border-slate-200 rounded-xl [&_.ql-toolbar]:border-none [&_.ql-toolbar]:bg-slate-50 [&_.ql-toolbar]:border-b [&_.ql-toolbar]:border-slate-200 [&_.ql-toolbar]:rounded-t-xl [&_.ql-container]:border-none [&_.ql-editor]:min-h-24">
                            <ReactQuill
                              theme="snow"
                              modules={quillModules}
                              value={configuracion.cabeceraGlobalHtml || ''}
                              onChange={(val) => actualizarConfiguracion('cabeceraGlobalHtml', val)}
                            />
                          </div>
                        )}
                      </div>

                      <div className="space-y-3 pt-4 border-t border-slate-100">
                        <label className="flex items-center gap-3 text-sm font-bold text-slate-700 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={!!configuracion.mostrarPiePaginaGlobal}
                            onChange={(e) => actualizarConfiguracion('mostrarPiePaginaGlobal', e.target.checked)}
                            className="w-4 h-4 rounded text-blue-600"
                          />
                          Habilitar Pie de Página Global en todas las páginas de contenido
                        </label>
                        {configuracion.mostrarPiePaginaGlobal && (
                          <div className="bg-white border-2 border-slate-200 rounded-xl [&_.ql-toolbar]:border-none [&_.ql-toolbar]:bg-slate-50 [&_.ql-toolbar]:border-b [&_.ql-toolbar]:border-slate-200 [&_.ql-toolbar]:rounded-t-xl [&_.ql-container]:border-none [&_.ql-editor]:min-h-24">
                            <ReactQuill
                              theme="snow"
                              modules={quillModules}
                              value={configuracion.piePaginaGlobalHtml || ''}
                              onChange={(val) => actualizarConfiguracion('piePaginaGlobalHtml', val)}
                            />
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                  
                  <div className="mt-16 pt-10 border-t-2 border-red-100 flex flex-col md:flex-row items-center justify-between gap-6 bg-red-50/50 p-8 rounded-3xl">
                    <div>
                      <h4 className="text-lg font-black text-red-700 flex items-center gap-2"><AlertCircle size={20} /> Zona de Eliminación</h4>
                    </div>
                    <button type="button" onClick={intentarEliminarEdicion} disabled={eliminando} className="shrink-0 flex items-center justify-center space-x-2 px-6 py-3.5 bg-white border-2 border-red-200 text-red-600 rounded-xl text-sm font-bold hover:bg-red-600 hover:text-white transition-all shadow-sm cursor-pointer">
                      <Trash2 size={18} /> <span>{eliminando ? 'Eliminando...' : 'Eliminar Edición'}</span>
                    </button>
                  </div>
                </form>
              )}

              {pestañaActiva === 'contenido' && paginaActiva && (
                <div className="animate-in fade-in duration-500">
                  
                  {/* BARRA SUPERIOR DE NAVEGACIÓN DE PÁGINAS CON SALTO DIRECTO, DUPLICACIÓN Y EXPORTACIÓN */}
                  <div className="bg-white p-4 rounded-2xl shadow-sm border border-slate-200 flex flex-wrap gap-3 justify-between items-center mb-8 sticky top-4 z-40">
                    <div className="flex flex-wrap items-center gap-2 w-full md:w-auto justify-center md:justify-start">
                      <button onClick={() => cambiarPagina(Math.max(0, indicePaginaActual - 1))} disabled={indicePaginaActual === 0} className="p-2 hover:bg-slate-100 rounded-lg disabled:opacity-30 text-slate-600 cursor-pointer"><ChevronLeft size={20} /></button>
                      
                      <select
                        value={indicePaginaActual}
                        onChange={(e) => cambiarPagina(Number(e.target.value))}
                        className="font-bold text-slate-700 px-3 border-2 border-slate-200 hover:border-blue-400 py-2 rounded-lg text-sm bg-white outline-none max-w-60 truncate cursor-pointer"
                        title="Saltar directamente a cualquier hoja"
                      >
                        {paginas.map((p, idx) => (
                          <option key={p.id} value={idx}>
                            Hoja {idx + 1} de {paginas.length} — {p.titulo || p.subtitulo || 'Sin título'}
                          </option>
                        ))}
                      </select>

                      <div className="flex items-center gap-1 bg-slate-50 border border-slate-200 px-2 py-1.5 rounded-lg" title="Escribe el número de hoja y presiona Enter">
                        <span className="text-[10px] font-black text-slate-400 uppercase">Ir a:</span>
                        <input
                          type="number"
                          min={1}
                          max={paginas.length}
                          placeholder={`${indicePaginaActual + 1}`}
                          onKeyDown={(e) => {
                            if (e.key === 'Enter') {
                              const val = Number((e.target as HTMLInputElement).value);
                              if (val >= 1 && val <= paginas.length) {
                                cambiarPagina(val - 1);
                                (e.target as HTMLInputElement).value = '';
                              }
                            }
                          }}
                          className="w-12 text-xs font-black text-blue-700 text-center bg-white border border-slate-200 rounded py-0.5 outline-none focus:border-blue-500"
                        />
                      </div>

                      <button onClick={() => cambiarPagina(Math.min(paginas.length - 1, indicePaginaActual + 1))} disabled={indicePaginaActual === paginas.length - 1} className="p-2 hover:bg-slate-100 rounded-lg disabled:opacity-30 text-slate-600 cursor-pointer"><ChevronRight size={20} /></button>
                      
                      <div className="h-8 w-px bg-slate-200 mx-1 hidden md:block"></div>
                      
                      <div className="flex items-center gap-1 bg-slate-50 border border-slate-200 p-1 rounded-lg">
                        <span className="text-[10px] font-bold text-slate-400 uppercase px-2 hidden lg:inline">Orden:</span>
                        <button onClick={() => moverPagina('izq')} disabled={indicePaginaActual === 0} className="p-1.5 hover:bg-white rounded text-slate-500 hover:text-blue-600 disabled:opacity-30 transition-colors shadow-sm cursor-pointer" title="Mover página a la izquierda"><MoveLeft size={14} /></button>
                        <button onClick={() => moverPagina('der')} disabled={indicePaginaActual === paginas.length - 1} className="p-1.5 hover:bg-white rounded text-slate-500 hover:text-blue-600 disabled:opacity-30 transition-colors shadow-sm cursor-pointer" title="Mover página a la derecha"><MoveRight size={14} /></button>
                      </div>

                      <div className="h-8 w-px bg-slate-200 mx-1 hidden md:block"></div>

                      <button onClick={agregarPagina} className="text-xs font-bold text-blue-600 hover:bg-blue-50 px-3 py-2 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer"><Plus size={15} /> Nueva Pág</button>
                      <button onClick={duplicarPaginaActual} className="text-xs font-bold text-indigo-600 hover:bg-indigo-50 px-3 py-2 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer" title="Duplicar esta página completa con sus bloques"><Copy size={15} /> Duplicar Pág</button>
                      <button 
                        onClick={() => setModalImportarJsonAbierto(true)} 
                        className="text-xs font-bold text-emerald-700 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 px-3 py-2 rounded-lg transition-all flex items-center gap-1.5 shadow-xs cursor-pointer"
                        title="Inyectar lote de páginas, tablas y gráficas desde JSON"
                      >
                        <Database size={14} /> Importar JSON
                      </button>
                      <button onClick={intentarEliminarPagina} className="text-xs font-bold text-red-600 hover:bg-red-50 px-3 py-2 rounded-lg transition-all flex items-center gap-1.5 cursor-pointer"><Trash2 size={15} /> Eliminar</button>
                    </div>

                    <div className="flex items-center gap-3 w-full md:w-auto justify-center">
                      <div className="flex items-center gap-2 text-xs font-bold uppercase tracking-widest px-3">
                        {estadoGuardadoVivo === 'guardado' && <><CheckCircle2 size={14} className="text-green-500" /> Guardado</>}
                        {estadoGuardadoVivo === 'escribiendo' && <><Edit3 size={14} className="text-blue-500" /> Cambios...</>}
                        {estadoGuardadoVivo === 'guardando' && <><RefreshCcw size={14} className="text-yellow-500 animate-spin" /> Nube</>}
                      </div>
                    </div>
                  </div>

                  {/* CABECERA CONFIGURABLE DE LA PÁGINA */}
                  <div className="mb-8 bg-white p-6 rounded-2xl shadow-sm border border-slate-200 space-y-6">
                    <div className="flex flex-col md:flex-row md:items-start justify-between gap-6">
                      <div className="w-full space-y-4">
                        <div>
                          <label className="block text-xs font-bold text-slate-400 uppercase tracking-widest mb-1">Título de la Página (Déjalo en blanco para ocultarlo)</label>
                          <input
                            type="text"
                            value={paginaActiva.titulo}
                            onChange={(e) => actualizarTituloPagina(e.target.value)}
                            placeholder="Ej: Capítulo 1. Introducción (O deja vacío si es continuación)"
                            className="text-3xl font-black tracking-tight text-slate-900 bg-transparent outline-none border-b-2 border-transparent focus:border-blue-500 w-full placeholder:text-slate-300 transition-colors"
                          />
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-400 uppercase tracking-widest mb-1">Subtítulo de la Página (Opcional)</label>
                          <input
                            type="text"
                            value={paginaActiva.subtitulo || ''}
                            onChange={(e) => actualizarCampoPagina('subtitulo', e.target.value)}
                            placeholder="Escribe un subtítulo o descripción secundaria para esta hoja..."
                            className="text-lg font-medium text-slate-600 bg-transparent outline-none border-b-2 border-transparent focus:border-blue-500 w-full placeholder:text-slate-300 transition-colors"
                          />
                        </div>
                      </div>
                      
                      <div className="shrink-0 flex items-center gap-3 bg-slate-50 p-3 rounded-xl border border-slate-200">
                        <input 
                          type="color" 
                          value={paginaActiva.colorFondo || '#fdfbf7'} 
                          onChange={(e) => actualizarColorFondoPagina(e.target.value)} 
                          className="w-10 h-10 cursor-pointer rounded-lg border-2"
                          title="Color de fondo de la página"
                        />
                        <div>
                          <span className="block text-sm font-bold text-slate-800">Fondo de Página</span>
                          <span className="block text-xs text-slate-500">{paginaActiva.colorFondo || '#fdfbf7'}</span>
                        </div>
                      </div>
                    </div>

                    <div className="pt-4 border-t border-slate-100 flex flex-wrap gap-6 items-center">
                      <label className="flex items-center gap-2.5 text-sm font-bold text-slate-700 cursor-pointer bg-slate-50 px-4 py-2.5 rounded-xl border border-slate-200 hover:border-blue-300 transition-colors">
                        <input 
                          type="checkbox" 
                          checked={paginaActiva.mostrarNumero !== false}
                          onChange={(e) => actualizarCampoPagina('mostrarNumero', e.target.checked)}
                          className="w-4 h-4 rounded text-blue-600 focus:ring-blue-500"
                        />
                        Mostrar número de índice en el encabezado (Ej: &quot;1.&quot;, &quot;2.&quot;)
                      </label>

                      <label className="flex items-center gap-2.5 text-sm font-bold text-slate-700 cursor-pointer bg-amber-50/70 px-4 py-2.5 rounded-xl border border-amber-200 hover:border-amber-400 transition-colors">
                        <input 
                          type="checkbox" 
                          checked={paginaActiva.posicion === 'antes_portada'}
                          onChange={(e) => actualizarCampoPagina('posicion', e.target.checked ? 'antes_portada' : 'despues_portada')}
                          className="w-4 h-4 rounded text-amber-600 focus:ring-amber-500"
                        />
                        Ubicar esta hoja ANTES de la Portada Principal (Pre-Portada)
                      </label>
                    </div>

                    <details className="pt-3 border-t border-slate-100 group">
                      <summary className="text-xs font-bold uppercase tracking-widest text-blue-600 cursor-pointer select-none flex items-center gap-2">
                        <span>+ Configurar Cabecera y Pie de Página exclusivos para esta hoja</span>
                      </summary>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mt-4">
                        <div>
                          <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Cabecera de esta Página</label>
                          <div className="bg-white border border-slate-200 rounded-xl [&_.ql-toolbar]:border-none [&_.ql-toolbar]:bg-slate-50 [&_.ql-toolbar]:rounded-t-xl [&_.ql-container]:border-none [&_.ql-editor]:min-h-20">
                            <ReactQuill
                              theme="snow"
                              modules={quillModules}
                              value={paginaActiva.cabeceraHtml || ''}
                              onChange={(val) => actualizarCampoPagina('cabeceraHtml', val)}
                            />
                          </div>
                        </div>
                        <div>
                          <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Pie de Página de esta Hoja</label>
                          <div className="bg-white border border-slate-200 rounded-xl [&_.ql-toolbar]:border-none [&_.ql-toolbar]:bg-slate-50 [&_.ql-toolbar]:rounded-t-xl [&_.ql-container]:border-none [&_.ql-editor]:min-h-20">
                            <ReactQuill
                              theme="snow"
                              modules={quillModules}
                              value={paginaActiva.piePaginaHtml || ''}
                              onChange={(val) => actualizarCampoPagina('piePaginaHtml', val)}
                            />
                          </div>
                        </div>
                      </div>
                    </details>
                  </div>
                  <div className="space-y-6">
                    {paginaActiva.bloques.length === 0 && (
                      <div className="border-4 border-dashed border-slate-200 rounded-3xl p-20 flex flex-col items-center justify-center text-slate-400 bg-white/50">
                        <Layout size={48} className="mb-4 text-slate-300" />
                        <h3 className="text-xl font-bold text-slate-600 mb-2">Esta página está vacía</h3>
                        <p className="mb-6">Agrega un bloque estructural para comenzar a construirla.</p>
                      </div>
                    )}

                    {paginaActiva.bloques.map((bloque, bIdx) => (
                      <div key={bloque.id} className={`bg-white rounded-2xl shadow-sm border group ${bloque.tipo === 'tabla' && bloque.oculto ? 'border-amber-300 opacity-90' : 'border-slate-200'}`}>
                        
                        <div className={`px-4 py-3 flex flex-wrap gap-2 justify-between items-center border-b rounded-t-2xl ${bloque.tipo === 'tabla' && bloque.oculto ? 'bg-amber-50 border-amber-200' : 'bg-slate-50 border-slate-100'}`}>
                          <div className="flex items-center gap-3 text-slate-400">
                            <GripVertical size={16} className="cursor-grab active:cursor-grabbing hover:text-slate-600 hidden md:block" />
                            <span className={`text-xs font-bold uppercase tracking-widest flex items-center gap-2 ${bloque.tipo === 'tabla' && bloque.oculto ? 'text-amber-700' : 'text-slate-500'}`}>
                              {getBloqueIcon(bloque)}
                              Bloque de {bloque.tipo} {bloque.tipo === 'tabla' && bloque.oculto && '(Oculto al público)'}
                            </span>
                          </div>
                          
                          <div className="flex items-center gap-3 flex-wrap">
                            {bloque.tipo === 'texto' && (
                              <div className="flex items-center gap-1.5 bg-white px-2.5 py-1 rounded-lg border border-slate-200 shadow-xs">
                                <Columns size={13} className="text-blue-600" />
                                <span className="text-[10px] font-black text-slate-500 uppercase">Diseño:</span>
                                <select
                                  value={bloque.columnas || 1}
                                  onChange={(e) => actualizarCampoBloqueTexto(bloque.id, 'columnas', Number(e.target.value) as 1 | 2 | 3)}
                                  className="text-xs font-bold text-blue-700 bg-transparent outline-none cursor-pointer"
                                >
                                  <option value={1}>1 Columna (Ancho completo)</option>
                                  <option value={2}>2 Columnas Paralelas (Lado a Lado)</option>
                                  <option value={3}>3 Columnas Paralelas</option>
                                </select>
                              </div>
                            )}

                            {bloque.tipo === 'tabla' && (
                              <>
                                <button
                                  type="button"
                                  onClick={() => clonarTablaExistente(bloque)}
                                  className="flex items-center gap-1.5 bg-white hover:bg-blue-50 text-slate-600 hover:text-blue-600 px-2.5 py-1 rounded-lg border border-slate-200 text-[11px] font-bold transition-colors shadow-xs cursor-pointer"
                                  title="Duplicar esta tabla en esta misma página"
                                >
                                  <Copy size={12} /> Clonar
                                </button>
                                <div className="flex items-center gap-1.5 bg-white px-2.5 py-1 rounded-lg border border-slate-200 shadow-xs" title="Reducir o ampliar el espacio vertical entre todas las filas (0px = ultra compacto)">
                                  <span className="text-[10px] font-black text-slate-500 uppercase">Espacio Filas:</span>
                                  <input
                                    type="range"
                                    min="0"
                                    max="24"
                                    step="1"
                                    value={bloque.paddingFila ?? 6}
                                    onChange={(e) => actualizarPropiedadTabla(bloque.id, 'paddingFila', Number(e.target.value))}
                                    className="w-16 accent-emerald-600 cursor-pointer"
                                  />
                                  <span className="text-[11px] font-black text-emerald-700 min-w-7 text-right">{bloque.paddingFila ?? 6}px</span>
                                </div>
                                <div className="flex items-center gap-1.5 bg-white px-2.5 py-1 rounded-lg border border-slate-200 shadow-xs" title="Tamaño de letra de la tabla">
                                  <span className="text-[10px] font-black text-slate-500 uppercase">Letra:</span>
                                  <input
                                    type="number"
                                    min="8"
                                    max="28"
                                    value={bloque.tamanoTextoTabla ?? 14}
                                    onChange={(e) => actualizarPropiedadTabla(bloque.id, 'tamanoTextoTabla', Number(e.target.value))}
                                    className="w-11 text-xs font-bold text-slate-700 outline-none text-center"
                                  />
                                  <span className="text-[10px] text-slate-400">px</span>
                                </div>
                              </>
                            )}

                            {(bloque.tipo === 'tabla' || bloque.tipo === 'grafica') && (
                              <div className="flex items-center gap-2 bg-white px-3 py-1 rounded-lg border border-slate-200 shadow-xs">
                                <span className="text-[10px] font-black text-slate-500 uppercase">Tamaño:</span>
                                <input
                                  type="range"
                                  min="30"
                                  max="200"
                                  step="5"
                                  value={bloque.escala || 100}
                                  onChange={(e) => actualizarEscalaBloque(bloque.id, Number(e.target.value))}
                                  className="w-20 accent-blue-600 cursor-pointer"
                                />
                                <span className="text-xs font-black text-blue-600 min-w-10 text-right">{bloque.escala || 100}%</span>
                              </div>
                            )}

                            <div className="flex items-center gap-1 bg-slate-200/50 p-1 rounded-lg">
                              <button onClick={() => moverBloque(bIdx, 'arriba')} disabled={bIdx === 0} className="p-1 hover:bg-white rounded text-slate-500 hover:text-blue-600 disabled:opacity-30 transition-colors shadow-sm cursor-pointer" title="Subir bloque"><ArrowUp size={14} /></button>
                              <button onClick={() => moverBloque(bIdx, 'abajo')} disabled={bIdx === paginaActiva.bloques.length - 1} className="p-1 hover:bg-white rounded text-slate-500 hover:text-blue-600 disabled:opacity-30 transition-colors shadow-sm cursor-pointer" title="Bajar bloque"><ArrowDown size={14} /></button>
                            </div>

                            <div className="w-px h-4 bg-slate-300"></div>

                            <div className="flex items-center gap-2 text-xs font-bold text-slate-400" title="Cambiar color de fondo general del bloque">
                              <Palette size={14} />
                              <input 
                                type="color" 
                                value={bloque.colorFondo || '#ffffff'} 
                                onChange={(e) => actualizarColorFondoBloque(bloque.id, e.target.value)} 
                                className="w-5 h-5 rounded cursor-pointer border-none p-0"
                              />
                            </div>
                            
                            <div className="w-px h-4 bg-slate-300"></div>
                            <button onClick={() => intentarEliminarBloque(bloque.id, bloque.tipo)} className="text-red-400 hover:text-red-600 transition-colors p-1 cursor-pointer"><Trash2 size={16} /></button>
                          </div>
                        </div>

                        <div className="p-0">
                          {bloque.tipo === 'texto' && (() => {
                            const colsActivas: (1 | 2 | 3)[] = bloque.columnas === 3 ? [1, 2, 3] : bloque.columnas === 2 ? [1, 2] : [1];
                            const todasImgsEnBloque = colsActivas.flatMap(colNum => {
                              const htmlCol = colNum === 2 ? bloque.contenidoCol2 : colNum === 3 ? bloque.contenidoCol3 : bloque.contenido;
                              return extraerImagenesDeColumnaTexto(htmlCol).map(imgObj => ({ ...imgObj, colNum }));
                            });

                            return (
                              <div>
                                {/* BARRA FLOTANTE AUTOMÁTICA CUANDO EL USUARIO SOMBREA TEXTO EN EL EDITOR */}
                                {seleccionTextoActiva && seleccionTextoActiva.bloqueId === bloque.id && (
                                  <div className="bg-slate-900 text-white px-4 py-3 flex flex-wrap items-center justify-between gap-3 border-b border-slate-700 animate-in fade-in duration-200">
                                    <div className="flex items-center gap-2 min-w-0">
                                      <span className="px-2 py-0.5 bg-blue-600 text-[10px] font-black uppercase rounded">Texto Seleccionado</span>
                                      <span className="text-xs font-bold text-blue-200 truncate max-w-xs">&ldquo;{seleccionTextoActiva.texto}&rdquo;</span>
                                    </div>
                                    <div className="flex flex-wrap items-center gap-2.5">
                                      <select id={`float-page-${bloque.id}`} defaultValue="" className="px-2.5 py-1.5 text-xs font-bold bg-slate-800 text-white border border-slate-700 rounded-lg outline-none max-w-52">
                                        <option value="">-- Elegir Página del Libro --</option>
                                        {paginas.map((p, idx) => (
                                          <option key={p.id} value={p.id}>Hoja {idx + 1}: {p.titulo || 'Sin título'}</option>
                                        ))}
                                      </select>
                                      <div className="flex items-center gap-1.5 bg-slate-800 px-2.5 py-1 rounded-lg border border-slate-700" title="Color del texto enlazado">
                                        <span className="text-[10px] font-bold text-slate-300">Color:</span>
                                        <input type="color" id={`float-color-${bloque.id}`} defaultValue={edicionActiva?.color_principal || '#1d4ed8'} className="w-5 h-5 rounded cursor-pointer border-0 p-0" />
                                      </div>
                                      <label className="flex items-center gap-1.5 text-xs font-bold text-slate-200 cursor-pointer bg-slate-800 px-2.5 py-1.5 rounded-lg border border-slate-700">
                                        <input type="checkbox" id={`float-under-${bloque.id}`} defaultChecked={true} className="w-3.5 h-3.5 rounded text-blue-500" />
                                        Subrayado
                                      </label>
                                      <button
                                        type="button"
                                        onClick={() => {
                                          const pagId = (document.getElementById(`float-page-${bloque.id}`) as HTMLSelectElement)?.value || '';
                                          const colHex = (document.getElementById(`float-color-${bloque.id}`) as HTMLInputElement)?.value || '#1d4ed8';
                                          const sub = (document.getElementById(`float-under-${bloque.id}`) as HTMLInputElement)?.checked ?? true;
                                          if (!pagId) {
                                            mostrarMensaje('Elige primero la página del libro en la lista desplegable.', 'error');
                                            return;
                                          }
                                          enlazarTextoEnColumna(bloque.id, seleccionTextoActiva.columna, '', 'pagina', pagId, colHex, sub);
                                        }}
                                        className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white text-xs font-black rounded-lg shadow transition-colors cursor-pointer"
                                      >
                                        Enlazar a Página
                                      </button>
                                      <input
                                        type="text"
                                        id={`float-url-${bloque.id}`}
                                        placeholder="O pega URL..."
                                        className="px-2.5 py-1.5 text-xs bg-slate-800 text-white border border-slate-700 rounded-lg outline-none w-32"
                                      />
                                      <button
                                        type="button"
                                        onClick={() => {
                                          const urlEx = (document.getElementById(`float-url-${bloque.id}`) as HTMLInputElement)?.value || '';
                                          const colHex = (document.getElementById(`float-color-${bloque.id}`) as HTMLInputElement)?.value || '#1d4ed8';
                                          const sub = (document.getElementById(`float-under-${bloque.id}`) as HTMLInputElement)?.checked ?? true;
                                          if (!urlEx.trim()) {
                                            mostrarMensaje('Escribe la URL externa.', 'error');
                                            return;
                                          }
                                          enlazarTextoEnColumna(bloque.id, seleccionTextoActiva.columna, '', 'url', urlEx, colHex, sub);
                                        }}
                                        className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-black rounded-lg shadow transition-colors cursor-pointer"
                                      >
                                        Enlazar URL
                                      </button>
                                      <button
                                        type="button"
                                        onClick={() => setSeleccionTextoActiva(null)}
                                        className="p-1 text-slate-400 hover:text-white cursor-pointer"
                                        title="Cerrar barra de selección"
                                      >
                                        <X size={15} />
                                      </button>
                                    </div>
                                  </div>
                                )}

                                {/* EDITORES DE TEXTO EN 1, 2 O 3 COLUMNAS INDEPENDIENTES LADO A LADO */}
                                <div 
                                  className={`grid grid-cols-1 ${bloque.columnas === 2 ? 'md:grid-cols-2 divide-y md:divide-y-0 md:divide-x divide-slate-200' : bloque.columnas === 3 ? 'md:grid-cols-3 divide-y md:divide-y-0 md:divide-x divide-slate-200' : ''}`}
                                  style={{ backgroundColor: bloque.colorFondo || 'transparent' }}
                                >
                                  <div 
                                    className="[&_.ql-toolbar]:border-none [&_.ql-toolbar]:bg-slate-50 [&_.ql-container]:border-none [&_.ql-editor]:min-h-37.5 [&_.ql-editor]:text-base"
                                    onMouseUp={(e) => capturarSeleccionEditor(bloque.id, 1, e)}
                                    onKeyUp={(e) => capturarSeleccionEditor(bloque.id, 1, e)}
                                  >
                                    {(bloque.columnas === 2 || bloque.columnas === 3) && (
                                      <div className="bg-slate-100/80 px-4 py-1.5 text-[10px] font-black uppercase tracking-widest text-slate-500 border-b border-slate-200">
                                        Columna 1 (Izquierda)
                                      </div>
                                    )}
                                    <ReactQuill theme="snow" modules={quillModules} value={bloque.contenido} onChange={(val) => actualizarBloqueTexto(bloque.id, val)} />
                                  </div>

                                  {(bloque.columnas === 2 || bloque.columnas === 3) && (
                                    <div 
                                      className="[&_.ql-toolbar]:border-none [&_.ql-toolbar]:bg-slate-50 [&_.ql-container]:border-none [&_.ql-editor]:min-h-37.5 [&_.ql-editor]:text-base"
                                      onMouseUp={(e) => capturarSeleccionEditor(bloque.id, 2, e)}
                                      onKeyUp={(e) => capturarSeleccionEditor(bloque.id, 2, e)}
                                    >
                                      <div className="bg-slate-100/80 px-4 py-1.5 text-[10px] font-black uppercase tracking-widest text-slate-500 border-b border-slate-200">
                                        Columna 2 ({bloque.columnas === 2 ? 'Derecha' : 'Centro'})
                                      </div>
                                      <ReactQuill theme="snow" modules={quillModules} value={bloque.contenidoCol2 || ''} onChange={(val) => actualizarCampoBloqueTexto(bloque.id, 'contenidoCol2', val)} />
                                    </div>
                                  )}

                                  {bloque.columnas === 3 && (
                                    <div 
                                      className="[&_.ql-toolbar]:border-none [&_.ql-toolbar]:bg-slate-50 [&_.ql-container]:border-none [&_.ql-editor]:min-h-37.5 [&_.ql-editor]:text-base"
                                      onMouseUp={(e) => capturarSeleccionEditor(bloque.id, 3, e)}
                                      onKeyUp={(e) => capturarSeleccionEditor(bloque.id, 3, e)}
                                    >
                                      <div className="bg-slate-100/80 px-4 py-1.5 text-[10px] font-black uppercase tracking-widest text-slate-500 border-b border-slate-200">
                                        Columna 3 (Derecha)
                                      </div>
                                      <ReactQuill theme="snow" modules={quillModules} value={bloque.contenidoCol3 || ''} onChange={(val) => actualizarCampoBloqueTexto(bloque.id, 'contenidoCol3', val)} />
                                    </div>
                                  )}
                                </div>

                                {/* BARRA DE HERRAMIENTAS AVANZADA DEL BLOQUE DE TEXTO */}
                                <div className="bg-slate-50 p-4 border-t border-slate-200 space-y-3 rounded-b-2xl">
                                  
                                  {/* 1. ENLAZAR TEXTO SELECCIONADO O LÍNEA DE ÍNDICE */}
                                  <div className="flex flex-wrap items-center gap-2.5 bg-white p-3 rounded-xl border border-slate-200">
                                    <span className="text-[11px] font-black text-blue-700 uppercase tracking-wider flex items-center gap-1.5">
                                      <Link2 size={14} /> Enlazar Texto (Sombréalo arriba):
                                    </span>
                                    {(bloque.columnas === 2 || bloque.columnas === 3) && (
                                      <select id={`link-col-${bloque.id}`} defaultValue="1" className="px-2 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg outline-none">
                                        <option value="1">Columna 1</option>
                                        <option value="2">Columna 2</option>
                                        {bloque.columnas === 3 && <option value="3">Columna 3</option>}
                                      </select>
                                    )}
                                    <input
                                      type="text"
                                      id={`link-txt-${bloque.id}`}
                                      placeholder="Sombrea texto arriba o escríbelo..."
                                      className="px-3 py-1.5 text-xs font-semibold bg-slate-50 border border-slate-200 rounded-lg outline-none focus:border-blue-500 min-w-48"
                                    />
                                    <select id={`link-page-${bloque.id}`} defaultValue="" className="px-2.5 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg outline-none max-w-52">
                                      <option value="">-- Elegir Página Destino --</option>
                                      {paginas.map((p, idx) => (
                                        <option key={p.id} value={p.id}>
                                          Hoja {idx + 1}: {p.titulo || 'Sin título'}
                                        </option>
                                      ))}
                                    </select>
                                    <div className="flex items-center gap-1.5 bg-slate-50 px-2.5 py-1 rounded-lg border border-slate-200" title="Elige el color del texto enlazado">
                                      <span className="text-[10px] font-bold text-slate-500">Color:</span>
                                      <input
                                        type="color"
                                        id={`link-color-${bloque.id}`}
                                        defaultValue={edicionActiva?.color_principal || '#1d4ed8'}
                                        className="w-5 h-5 rounded cursor-pointer border-0 p-0"
                                      />
                                    </div>
                                    <label className="flex items-center gap-1.5 text-xs font-bold text-slate-600 cursor-pointer bg-slate-50 px-2.5 py-1.5 rounded-lg border border-slate-200">
                                      <input
                                        type="checkbox"
                                        id={`link-under-${bloque.id}`}
                                        defaultChecked={true}
                                        className="w-3.5 h-3.5 rounded text-blue-600"
                                      />
                                      Subrayado
                                    </label>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        const col = Number((document.getElementById(`link-col-${bloque.id}`) as HTMLSelectElement)?.value || 1) as 1 | 2 | 3;
                                        const txt = (document.getElementById(`link-txt-${bloque.id}`) as HTMLInputElement)?.value || '';
                                        const pagId = (document.getElementById(`link-page-${bloque.id}`) as HTMLSelectElement)?.value || '';
                                        const colHex = (document.getElementById(`link-color-${bloque.id}`) as HTMLInputElement)?.value || '#1d4ed8';
                                        const sub = (document.getElementById(`link-under-${bloque.id}`) as HTMLInputElement)?.checked ?? true;
                                        if (!pagId) {
                                          mostrarMensaje('Selecciona primero la página del libro a la que deseas redirigir.', 'error');
                                          return;
                                        }
                                        enlazarTextoEnColumna(bloque.id, col, txt, 'pagina', pagId, colHex, sub);
                                      }}
                                      className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg transition-colors shadow-xs cursor-pointer"
                                    >
                                      Enlazar a Página
                                    </button>
                                    <span className="text-[10px] font-bold text-slate-400">o URL:</span>
                                    <input
                                      type="text"
                                      id={`link-url-${bloque.id}`}
                                      placeholder="https://..."
                                      className="px-2.5 py-1.5 text-xs bg-slate-50 border border-slate-200 rounded-lg outline-none w-32"
                                    />
                                    <button
                                      type="button"
                                      onClick={() => {
                                        const col = Number((document.getElementById(`link-col-${bloque.id}`) as HTMLSelectElement)?.value || 1) as 1 | 2 | 3;
                                        const txt = (document.getElementById(`link-txt-${bloque.id}`) as HTMLInputElement)?.value || '';
                                        const urlEx = (document.getElementById(`link-url-${bloque.id}`) as HTMLInputElement)?.value || '';
                                        const colHex = (document.getElementById(`link-color-${bloque.id}`) as HTMLInputElement)?.value || '#1d4ed8';
                                        const sub = (document.getElementById(`link-under-${bloque.id}`) as HTMLInputElement)?.checked ?? true;
                                        if (!urlEx.trim()) {
                                          mostrarMensaje('Escribe la URL externa.', 'error');
                                          return;
                                        }
                                        enlazarTextoEnColumna(bloque.id, col, txt, 'url', urlEx, colHex, sub);
                                      }}
                                      className="px-3 py-1.5 bg-slate-800 hover:bg-slate-900 text-white text-xs font-bold rounded-lg transition-colors cursor-pointer"
                                    >
                                      Enlazar URL
                                    </button>
                                  </div>

                                  {/* 2. INSERTAR Y GESTIONAR IMÁGENES EN EL TEXTO (EN POSICIÓN DEL CURSOR, BORDES, ALINEACIÓN, MOVER Y BORRAR) */}
                                  <div className="bg-white p-3.5 rounded-xl border border-slate-200 space-y-3">
                                    <div className="flex flex-wrap items-center gap-2.5">
                                      <span className="text-[11px] font-black text-pink-600 uppercase tracking-wider flex items-center gap-1.5">
                                        <ImageIcon size={14} /> Insertar Imagen (Haz clic donde quieras en el texto y súbela):
                                      </span>
                                      {(bloque.columnas === 2 || bloque.columnas === 3) && (
                                        <select id={`img-col-${bloque.id}`} defaultValue="1" className="px-2 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg outline-none">
                                          <option value="1">Columna 1</option>
                                          <option value="2">Columna 2</option>
                                          {bloque.columnas === 3 && <option value="3">Columna 3</option>}
                                        </select>
                                      )}
                                      <select id={`img-size-${bloque.id}`} defaultValue="50%" className="px-2.5 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg outline-none">
                                        <option value="15%">Ancho: 15% (Icono / Logo)</option>
                                        <option value="25%">Ancho: 25% (Miniatura)</option>
                                        <option value="35%">Ancho: 35% (Pequeña)</option>
                                        <option value="50%">Ancho: 50% (Mediana)</option>
                                        <option value="75%">Ancho: 75% (Grande)</option>
                                        <option value="100%">Ancho: 100% (Completa)</option>
                                      </select>
                                      <select id={`img-align-${bloque.id}`} defaultValue="centro" className="px-2.5 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg outline-none">
                                        <option value="izquierda">Alinear Izquierda</option>
                                        <option value="centro">Centrada</option>
                                        <option value="derecha">Alinear Derecha</option>
                                      </select>
                                      <select id={`img-border-${bloque.id}`} defaultValue="cuadrado" className="px-2.5 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg outline-none">
                                        <option value="cuadrado">Esquinas: Cuadradas (Rectas)</option>
                                        <option value="redondeado">Esquinas: Redondeadas</option>
                                        <option value="circular">Esquinas: Circular / Óvalo</option>
                                      </select>
                                      <div className="relative inline-block">
                                        <input
                                          type="file"
                                          accept="image/*"
                                          onChange={(e) => {
                                            const col = Number((document.getElementById(`img-col-${bloque.id}`) as HTMLSelectElement)?.value || 1) as 1 | 2 | 3;
                                            const sz = (document.getElementById(`img-size-${bloque.id}`) as HTMLSelectElement)?.value || '50%';
                                            const al = (document.getElementById(`img-align-${bloque.id}`) as HTMLSelectElement)?.value || 'centro';
                                            const brd = ((document.getElementById(`img-border-${bloque.id}`) as HTMLSelectElement)?.value || 'cuadrado') as 'cuadrado' | 'redondeado' | 'circular';
                                            manejarSubidaImagenEnTexto(e, bloque.id, col, sz, al, brd);
                                          }}
                                          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                                        />
                                        <button type="button" className="px-3.5 py-1.5 bg-pink-600 hover:bg-pink-700 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-xs cursor-pointer">
                                          <Upload size={12} /> Subir Imagen en Posición del Cursor
                                        </button>
                                      </div>
                                    </div>

                                    {/* GESTOR INDIVIDUAL DE IMÁGENES YA INSERTADAS EN ESTE BLOQUE DE TEXTO */}
                                    {todasImgsEnBloque.length > 0 && (
                                      <div className="pt-3 border-t border-slate-100 space-y-2">
                                        <span className="block text-[10px] font-black uppercase tracking-widest text-slate-400">
                                          Imágenes dentro de este texto ({todasImgsEnBloque.length}) — Edita posición, alineación, esquinas, tamaño o elimínalas:
                                        </span>
                                        <div className="grid grid-cols-1 gap-2">
                                          {todasImgsEnBloque.map((imgItem, iGlobal) => (
                                            <div key={`${imgItem.colNum}-${imgItem.indice}-${iGlobal}`} className="flex flex-wrap items-center justify-between gap-2 bg-slate-50 p-2.5 rounded-xl border border-slate-200">
                                              <div className="flex items-center gap-3 min-w-0">
                                                <img src={imgItem.src} alt="Miniatura" className="w-11 h-11 object-cover rounded border border-slate-300 bg-white shrink-0" />
                                                <div className="text-xs">
                                                  <span className="font-black text-slate-700 block">Imagen #{imgItem.indice + 1} {bloque.columnas && bloque.columnas > 1 ? `(Col ${imgItem.colNum})` : ''}</span>
                                                  <span className="text-[10px] text-slate-400">Ancho actual: {imgItem.ancho}</span>
                                                </div>
                                              </div>

                                              <div className="flex flex-wrap items-center gap-1.5">
                                                {/* Mover arriba o abajo entre párrafos */}
                                                <div className="flex items-center bg-white border border-slate-200 rounded-lg p-0.5" title="Subir o bajar la imagen entre los párrafos del texto">
                                                  <button
                                                    type="button"
                                                    onClick={() => moverImagenEntreParrafosTexto(bloque.id, imgItem.colNum, imgItem.indice, 'arriba')}
                                                    className="px-2 py-1 text-[11px] font-bold text-slate-600 hover:text-blue-600 hover:bg-blue-50 rounded flex items-center gap-0.5 cursor-pointer"
                                                    title="Subir imagen un párrafo hacia arriba"
                                                  >
                                                    <ArrowUp size={12} /> Subir
                                                  </button>
                                                  <button
                                                    type="button"
                                                    onClick={() => moverImagenEntreParrafosTexto(bloque.id, imgItem.colNum, imgItem.indice, 'abajo')}
                                                    className="px-2 py-1 text-[11px] font-bold text-slate-600 hover:text-blue-600 hover:bg-blue-50 rounded flex items-center gap-0.5 cursor-pointer"
                                                    title="Bajar imagen un párrafo hacia abajo"
                                                  >
                                                    <ArrowDown size={12} /> Bajar
                                                  </button>
                                                </div>

                                                {/* Alineación Izquierda / Centro / Derecha */}
                                                <div className="flex items-center bg-white border border-slate-200 rounded-lg p-0.5">
                                                  <button
                                                    type="button"
                                                    onClick={() => actualizarImagenEspecificaEnTexto(bloque.id, imgItem.colNum, imgItem.indice, { alineacion: 'izquierda' })}
                                                    className={`p-1.5 rounded cursor-pointer ${imgItem.alineacion === 'izquierda' ? 'bg-blue-600 text-white' : 'text-slate-600 hover:bg-slate-100'}`}
                                                    title="Alinear a la izquierda"
                                                  >
                                                    <AlignLeft size={13} />
                                                  </button>
                                                  <button
                                                    type="button"
                                                    onClick={() => actualizarImagenEspecificaEnTexto(bloque.id, imgItem.colNum, imgItem.indice, { alineacion: 'centro' })}
                                                    className={`p-1.5 rounded cursor-pointer ${imgItem.alineacion === 'centro' ? 'bg-blue-600 text-white' : 'text-slate-600 hover:bg-slate-100'}`}
                                                    title="Centrar imagen"
                                                  >
                                                    <AlignCenter size={13} />
                                                  </button>
                                                  <button
                                                    type="button"
                                                    onClick={() => actualizarImagenEspecificaEnTexto(bloque.id, imgItem.colNum, imgItem.indice, { alineacion: 'derecha' })}
                                                    className={`p-1.5 rounded cursor-pointer ${imgItem.alineacion === 'derecha' ? 'bg-blue-600 text-white' : 'text-slate-600 hover:bg-slate-100'}`}
                                                    title="Alinear a la derecha"
                                                  >
                                                    <AlignRight size={13} />
                                                  </button>
                                                </div>

                                                {/* Esquinas Cuadradas vs Redondeadas */}
                                                <select
                                                  value={imgItem.borde}
                                                  onChange={(e) => actualizarImagenEspecificaEnTexto(bloque.id, imgItem.colNum, imgItem.indice, { borde: e.target.value as 'cuadrado' | 'redondeado' | 'circular' })}
                                                  className="px-2 py-1 text-xs font-bold bg-white border border-slate-200 rounded-lg outline-none cursor-pointer"
                                                  title="Cambiar estilo de esquinas de esta imagen"
                                                >
                                                  <option value="cuadrado">Recta (Sin bordes redondos)</option>
                                                  <option value="redondeado">Esquinas Redondeadas</option>
                                                  <option value="circular">Circular</option>
                                                </select>

                                                {/* Tamaño / Ancho en vivo */}
                                                <select
                                                  value={['15%', '25%', '35%', '50%', '75%', '100%'].includes(imgItem.ancho) ? imgItem.ancho : '50%'}
                                                  onChange={(e) => actualizarImagenEspecificaEnTexto(bloque.id, imgItem.colNum, imgItem.indice, { ancho: e.target.value })}
                                                  className="px-2 py-1 text-xs font-bold bg-white border border-slate-200 rounded-lg outline-none cursor-pointer"
                                                  title="Cambiar tamaño de esta imagen"
                                                >
                                                  <option value="15%">15%</option>
                                                  <option value="25%">25%</option>
                                                  <option value="35%">35%</option>
                                                  <option value="50%">50%</option>
                                                  <option value="75%">75%</option>
                                                  <option value="100%">100%</option>
                                                </select>

                                                {/* Botón Eliminar esta imagen */}
                                                <button
                                                  type="button"
                                                  onClick={() => eliminarImagenDeColumnaTexto(bloque.id, imgItem.colNum, imgItem.indice)}
                                                  className="px-2.5 py-1 bg-red-50 hover:bg-red-600 text-red-600 hover:text-white border border-red-200 rounded-lg text-xs font-bold flex items-center gap-1 transition-colors cursor-pointer"
                                                  title="Eliminar esta imagen del texto"
                                                >
                                                  <Trash2 size={12} /> Quitar
                                                </button>
                                              </div>
                                            </div>
                                          ))}
                                        </div>
                                      </div>
                                    )}
                                  </div>

                                  {/* 3. GENERADOR DE BOTONES INLINE DENTRO DEL TEXTO */}
                                  <div className="flex flex-wrap items-center gap-2.5 bg-white p-3 rounded-xl border border-slate-200">
                                    <span className="text-[11px] font-black text-slate-600 uppercase tracking-wider flex items-center gap-1.5">
                                      <MousePointerClick size={14} className="text-blue-600" /> Botón en Texto:
                                    </span>
                                    <input type="text" id={`in-btn-txt-${bloque.id}`} placeholder="Texto del botón" className="px-3 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg outline-none focus:border-blue-500" />
                                    <input type="text" id={`in-btn-url-${bloque.id}`} placeholder="https://... o elige página" className="px-3 py-1.5 text-xs font-semibold bg-slate-50 border border-slate-200 rounded-lg outline-none focus:border-blue-500" />
                                    <select
                                      onChange={(e) => {
                                        const inputUrl = document.getElementById(`in-btn-url-${bloque.id}`) as HTMLInputElement;
                                        if (inputUrl && e.target.value) inputUrl.value = `#pagina:${e.target.value}`;
                                      }}
                                      defaultValue=""
                                      className="px-2 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg outline-none max-w-44"
                                    >
                                      <option value="">Ir a página...</option>
                                      {paginas.map((p, idx) => (
                                        <option key={p.id} value={p.id}>Hoja {idx + 1}: {p.titulo || 'Sin título'}</option>
                                      ))}
                                    </select>
                                    <select id={`in-btn-style-${bloque.id}`} defaultValue="solido" className="px-2.5 py-1.5 text-xs font-bold bg-slate-50 border border-slate-200 rounded-lg outline-none">
                                      <option value="solido">Sólido</option>
                                      <option value="pildora">Píldora (Ovalado)</option>
                                      <option value="contorno">Contorno</option>
                                      <option value="3d">3D Relieve</option>
                                      <option value="4d">4D Glow (Iluminado)</option>
                                    </select>
                                    <div className="flex items-center gap-1 bg-slate-50 px-2 py-1 rounded-lg border border-slate-200" title="Color principal del botón">
                                      <span className="text-[10px] font-bold text-slate-500">Fondo:</span>
                                      <input type="color" id={`in-btn-bg-${bloque.id}`} defaultValue={edicionActiva?.color_principal || '#1d4ed8'} className="w-6 h-6 rounded cursor-pointer border-0 p-0" />
                                    </div>
                                    <div className="flex items-center gap-1 bg-slate-50 px-2 py-1 rounded-lg border border-slate-200" title="Color del texto">
                                      <span className="text-[10px] font-bold text-slate-500">Texto:</span>
                                      <input type="color" id={`in-btn-tc-${bloque.id}`} defaultValue="#ffffff" className="w-6 h-6 rounded cursor-pointer border-0 p-0" />
                                    </div>
                                    <div className="flex items-center gap-1 bg-slate-50 px-2 py-1 rounded-lg border border-blue-200" title="Color de iluminación (para 4D Glow)">
                                      <span className="text-[10px] font-bold text-blue-600">Luz 4D:</span>
                                      <input type="color" id={`in-btn-glow-${bloque.id}`} defaultValue="#60a5fa" className="w-6 h-6 rounded cursor-pointer border-0 p-0" />
                                    </div>
                                    <button
                                      type="button"
                                      onClick={() => {
                                        const txt = (document.getElementById(`in-btn-txt-${bloque.id}`) as HTMLInputElement)?.value || 'Ver más';
                                        const url = (document.getElementById(`in-btn-url-${bloque.id}`) as HTMLInputElement)?.value || 'https://acofi.edu.co';
                                        const est = (document.getElementById(`in-btn-style-${bloque.id}`) as HTMLSelectElement)?.value || 'solido';
                                        const bg = (document.getElementById(`in-btn-bg-${bloque.id}`) as HTMLInputElement)?.value || '#1d4ed8';
                                        const tc = (document.getElementById(`in-btn-tc-${bloque.id}`) as HTMLInputElement)?.value || '#ffffff';
                                        const glow = (document.getElementById(`in-btn-glow-${bloque.id}`) as HTMLInputElement)?.value || '#60a5fa';
                                        const tag = `[[BOTON|${txt}|${url}|${est}|${bg}|${tc}|${glow}]]`;
                                        actualizarBloqueTexto(bloque.id, `${bloque.contenido}<p>${tag}</p>`);
                                      }}
                                      className="px-3.5 py-1.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg transition-colors shadow-sm cursor-pointer"
                                    >
                                      + Insertar Botón
                                    </button>
                                  </div>
                                </div>

                                {bloque.contenido.includes('[[BOTON|') && (
                                  <div className="p-5 bg-white border-t border-dashed border-slate-200 rounded-b-2xl">
                                    <span className="block text-[10px] font-bold uppercase tracking-widest text-slate-400 mb-3">Vista Previa Real del Texto con Botones:</span>
                                    <div className="prose prose-sm max-w-none" dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(bloque.contenido) }} />
                                  </div>
                                )}
                              </div>
                            );
                          })()}

                          {/* BLOQUE ESPECIAL DE FIRMAS CON CONTROL DE TAMAÑO REAL, POSICIÓN VERTICAL Y BOTÓN DE QUITAR SOLO IMAGEN */}
                          {bloque.tipo === 'firma' && (
                            <div className="p-6 space-y-6 rounded-b-2xl" style={{ backgroundColor: bloque.colorFondo || '#ffffff' }}>
                              <div className="flex flex-wrap items-center justify-between gap-4 bg-slate-50 p-4 rounded-xl border border-slate-200">
                                <div className="flex flex-wrap items-center gap-4">
                                  <div className="flex items-center gap-2">
                                    <span className="text-xs font-bold uppercase tracking-widest text-slate-500">Alineación:</span>
                                    <select
                                      value={bloque.alineacion || 'espaciado'}
                                      onChange={(e) => actualizarAlineacionBloqueFirma(bloque.id, e.target.value as BloqueFirma['alineacion'])}
                                      className="px-3 py-1.5 bg-white border border-slate-200 rounded-lg text-xs font-bold outline-none"
                                    >
                                      <option value="izquierda">Izquierda</option>
                                      <option value="centro">Centro</option>
                                      <option value="derecha">Derecha</option>
                                      <option value="espaciado">Simétrico (Lado a Lado)</option>
                                    </select>
                                  </div>

                                  {/* Control de espacio general superior e inferior de todo el bloque de firmas */}
                                  <div className="flex items-center gap-2 bg-white px-3 py-1.5 rounded-lg border border-slate-200">
                                    <span className="text-[11px] font-bold text-slate-600">Espacio Arriba del Bloque:</span>
                                    <input
                                      type="range"
                                      min="-40"
                                      max="140"
                                      step="4"
                                      value={bloque.espacioGeneralSuperior ?? 0}
                                      onChange={(e) => actualizarPropiedadBloqueFirma(bloque.id, 'espacioGeneralSuperior', Number(e.target.value))}
                                      className="w-20 accent-amber-600 cursor-pointer"
                                    />
                                    <span className="text-xs font-black text-amber-700 min-w-9 text-right">{bloque.espacioGeneralSuperior ?? 0}px</span>
                                  </div>
                                </div>

                                <button
                                  type="button"
                                  onClick={() => agregarSegundaFirmaAlBloque(bloque.id)}
                                  className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white text-xs font-bold rounded-xl flex items-center gap-1.5 shadow-xs transition-colors cursor-pointer"
                                >
                                  <Plus size={14} /> Agregar otra firma al lado
                                </button>
                              </div>

                              <div className={`grid grid-cols-1 ${bloque.firmas.length > 1 ? 'md:grid-cols-2' : ''} gap-6`}>
                                {bloque.firmas.map((fItem, fIdx) => {
                                  const keyLoader = `${bloque.id}_firma_${fIdx}`;
                                  return (
                                    <div key={fItem.id || fIdx} className="p-5 bg-slate-50 rounded-2xl border border-slate-200 space-y-4 relative">
                                      <div className="flex justify-between items-center">
                                        <span className="text-xs font-black uppercase tracking-widest text-amber-700">Firma #{fIdx + 1}</span>
                                        <div className="flex items-center gap-2">
                                          {fItem.url && (
                                            <button
                                              type="button"
                                              onClick={() => quitarSoloImagenFirma(bloque.id, fIdx)}
                                              className="bg-amber-100 hover:bg-amber-200 text-amber-900 px-2.5 py-1 rounded-lg text-[11px] font-bold flex items-center gap-1 transition-colors cursor-pointer"
                                              title="Borra únicamente la imagen de la firma pero conserva el nombre y cargo"
                                            >
                                              <X size={12} /> Quitar solo imagen
                                            </button>
                                          )}
                                          {bloque.firmas.length > 1 && (
                                            <button
                                              type="button"
                                              onClick={() => eliminarFirmaDelBloque(bloque.id, fIdx)}
                                              className="text-red-500 hover:bg-red-50 px-2 py-1 rounded-lg text-xs font-bold flex items-center gap-1 cursor-pointer"
                                            >
                                              <Trash2 size={13} /> Borrar firmante
                                            </button>
                                          )}
                                        </div>
                                      </div>

                                      <div className="flex flex-col sm:flex-row gap-3">
                                        <div className="relative flex-1">
                                          <input
                                            type="file"
                                            accept="image/*"
                                            onChange={(e) => manejarSubidaImagenFirma(e, bloque.id, fIdx)}
                                            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                                            disabled={subiendoImagen[keyLoader]}
                                          />
                                          <button type="button" className="w-full px-3 py-2.5 bg-white border border-slate-300 hover:border-blue-400 rounded-xl text-xs font-bold text-slate-700 flex items-center justify-center gap-2">
                                            {subiendoImagen[keyLoader] ? <RefreshCcw size={14} className="animate-spin" /> : <Upload size={14} />}
                                            {subiendoImagen[keyLoader] ? 'Subiendo firma...' : 'Subir Imagen de Firma (PNG/JPG)'}
                                          </button>
                                        </div>
                                        <input
                                          type="text"
                                          value={fItem.url || ''}
                                          onChange={(e) => actualizarItemFirma(bloque.id, fIdx, 'url', e.target.value)}
                                          placeholder="O pegar URL de la firma..."
                                          className="flex-1 px-3 py-2 bg-white border border-slate-200 rounded-xl text-xs outline-none focus:border-blue-500"
                                        />
                                      </div>

                                      {/* SLIDER DE TAMAÑO / ALTURA DE LA IMAGEN DE FIRMA (FUNCIONAL AL 100%) */}
                                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 bg-white p-3.5 rounded-xl border border-slate-200">
                                        <div>
                                          <label className="flex justify-between text-[11px] font-bold text-slate-600 uppercase mb-1">
                                            <span>Tamaño de Imagen Firma</span>
                                            <span className="text-blue-600 font-black">{fItem.tamanoFirma ?? 72}px</span>
                                          </label>
                                          <input
                                            type="range"
                                            min="16"
                                            max="240"
                                            step="2"
                                            value={fItem.tamanoFirma ?? 72}
                                            onChange={(e) => actualizarItemFirma(bloque.id, fIdx, 'tamanoFirma', Number(e.target.value))}
                                            className="w-full accent-blue-600 cursor-pointer"
                                          />
                                        </div>

                                        {/* SLIDER PARA SUBIR O BAJAR ESTA FIRMA INDIVIDUALMENTE */}
                                        <div>
                                          <label className="flex justify-between text-[11px] font-bold text-slate-600 uppercase mb-1">
                                            <span>Subir / Bajar Firma</span>
                                            <span className="text-amber-600 font-black">{fItem.espacioSuperior ?? 0}px</span>
                                          </label>
                                          <input
                                            type="range"
                                            min="-50"
                                            max="120"
                                            step="2"
                                            value={fItem.espacioSuperior ?? 0}
                                            onChange={(e) => actualizarItemFirma(bloque.id, fIdx, 'espacioSuperior', Number(e.target.value))}
                                            className="w-full accent-amber-600 cursor-pointer"
                                          />
                                        </div>
                                      </div>

                                      <div className="grid grid-cols-1 gap-3">
                                        <input
                                          type="text"
                                          value={fItem.nombre}
                                          onChange={(e) => actualizarItemFirma(bloque.id, fIdx, 'nombre', e.target.value)}
                                          placeholder="Nombre completo (Ej: Luis Alberto González Araujo)"
                                          className="w-full px-3.5 py-2.5 bg-white border border-slate-200 rounded-xl text-sm font-bold outline-none focus:border-blue-500"
                                        />
                                        <input
                                          type="text"
                                          value={fItem.cargo}
                                          onChange={(e) => actualizarItemFirma(bloque.id, fIdx, 'cargo', e.target.value)}
                                          placeholder="Cargo (Ej: Director ejecutivo ACOFI)"
                                          className="w-full px-3.5 py-2 bg-white border border-slate-200 rounded-xl text-xs font-semibold outline-none focus:border-blue-500"
                                        />
                                        <input
                                          type="text"
                                          value={fItem.infoAdicional || ''}
                                          onChange={(e) => actualizarItemFirma(bloque.id, fIdx, 'infoAdicional', e.target.value)}
                                          placeholder="Línea adicional opcional (Ej: Editor de la publicación)"
                                          className="w-full px-3.5 py-2 bg-white border border-slate-200 rounded-xl text-xs outline-none focus:border-blue-500"
                                        />
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>

                              {/* VISTA PREVIA FIEL DE LAS FIRMAS CON TAMAÑO Y ESPACIADO EN TIEMPO REAL */}
                              <div 
                                className={`p-6 bg-white rounded-2xl border border-dashed border-slate-300 flex flex-wrap gap-12 items-end ${
                                  bloque.alineacion === 'centro' ? 'justify-center text-center' :
                                  bloque.alineacion === 'derecha' ? 'justify-end text-right' :
                                  bloque.alineacion === 'espaciado' ? 'justify-around text-left' :
                                  'justify-start text-left'
                                }`}
                                style={{
                                  paddingTop: `${Math.max(8, 24 + (bloque.espacioGeneralSuperior ?? 0))}px`,
                                  paddingBottom: `${Math.max(8, 24 + (bloque.espacioGeneralInferior ?? 0))}px`
                                }}
                              >
                                {bloque.firmas.map((fItem, i) => (
                                  <div 
                                    key={i} 
                                    className="flex flex-col space-y-1 min-w-52 transition-all"
                                    style={{ marginTop: `${fItem.espacioSuperior ?? 0}px` }}
                                  >
                                    {fItem.url ? (
                                      <div className={`mb-2 flex ${bloque.alineacion === 'centro' ? 'justify-center' : bloque.alineacion === 'derecha' ? 'justify-end' : 'justify-start'}`}>
                                        <img
                                          src={fItem.url}
                                          alt={fItem.nombre}
                                          style={{ 
                                            height: `${fItem.tamanoFirma ?? 72}px`, 
                                            maxHeight: `${fItem.tamanoFirma ?? 72}px`,
                                            width: 'auto' 
                                          }}
                                          className="object-contain transition-all"
                                        />
                                      </div>
                                    ) : (
                                      <div className="h-10 w-44 border-b border-slate-300 mb-2 flex items-end pb-1 text-[10px] italic text-slate-400">
                                        (Sin imagen de firma — solo texto)
                                      </div>
                                    )}
                                    <span className="font-bold text-slate-900 text-base">{fItem.nombre}</span>
                                    <span className="text-slate-600 text-sm">{fItem.cargo}</span>
                                    {fItem.infoAdicional && <span className="text-slate-500 text-xs italic">{fItem.infoAdicional}</span>}
                                  </div>
                                ))}
                              </div>
                            </div>
                          )}

                          {/* CONFIGURADOR DE BLOQUE BOTÓN */}
                          {bloque.tipo === 'boton' && (
                            <div className="p-6 space-y-6 rounded-b-2xl" style={{ backgroundColor: bloque.colorFondo || '#ffffff' }}>
                              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 bg-slate-50 p-6 rounded-xl border border-slate-200">
                                <div>
                                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Texto del Botón</label>
                                  <input
                                    type="text"
                                    value={bloque.texto}
                                    onChange={(e) => actualizarBloqueBoton(bloque.id, 'texto', e.target.value)}
                                    className="w-full p-3 bg-white border border-slate-200 rounded-xl text-sm font-bold outline-none focus:border-blue-500"
                                    placeholder="Ej: Descargar Informe PDF"
                                  />
                                </div>
                                <div>
                                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Enlace / URL o Página del Libro</label>
                                  <div className="flex gap-2">
                                    <input
                                      type="text"
                                      value={bloque.url}
                                      onChange={(e) => actualizarBloqueBoton(bloque.id, 'url', e.target.value)}
                                      className="w-full p-3 bg-white border border-slate-200 rounded-xl text-sm font-semibold outline-none focus:border-blue-500"
                                      placeholder="https://... o selecciona una hoja"
                                    />
                                    <select
                                      value={bloque.url.startsWith('#pagina:') ? bloque.url.replace('#pagina:', '') : ''}
                                      onChange={(e) => {
                                        if (e.target.value) actualizarBloqueBoton(bloque.id, 'url', `#pagina:${e.target.value}`);
                                      }}
                                      className="p-3 bg-white border border-slate-200 rounded-xl text-xs font-bold outline-none max-w-44"
                                    >
                                      <option value="">O ir a página...</option>
                                      {paginas.map((p, idx) => (
                                        <option key={p.id} value={p.id}>Hoja {idx + 1}: {p.titulo || 'Sin título'}</option>
                                      ))}
                                    </select>
                                  </div>
                                </div>

                                <div>
                                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Estilo Visual del Botón</label>
                                  <select
                                    value={bloque.estilo}
                                    onChange={(e) => actualizarBloqueBoton(bloque.id, 'estilo', e.target.value as EstiloBoton)}
                                    className="w-full p-3 bg-white border border-slate-200 rounded-xl text-sm font-semibold outline-none focus:border-blue-500"
                                  >
                                    <option value="solido">Sólido Clásico (Bordes redondeados)</option>
                                    <option value="pildora">Ovalado / Píldora Moderna</option>
                                    <option value="contorno">Contorno Elegante (Transparente)</option>
                                    <option value="3d">3D con Relieve Físico</option>
                                    <option value="4d">4D Iluminado (Sombra de Neón Estática)</option>
                                  </select>
                                </div>

                                <div>
                                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Alineación en la Página</label>
                                  <select
                                    value={bloque.alineacion}
                                    onChange={(e) => actualizarBloqueBoton(bloque.id, 'alineacion', e.target.value)}
                                    className="w-full p-3 bg-white border border-slate-200 rounded-xl text-sm font-semibold outline-none focus:border-blue-500"
                                  >
                                    <option value="izquierda">Izquierda</option>
                                    <option value="centro">Centro</option>
                                    <option value="derecha">Derecha</option>
                                  </select>
                                </div>

                                <div className="col-span-1 md:col-span-2 grid grid-cols-1 md:grid-cols-3 gap-4 pt-2 border-t border-slate-200">
                                  <div className="flex items-center gap-3 bg-white p-3 rounded-xl border border-slate-200">
                                    <input
                                      type="color"
                                      value={bloque.colorBoton || '#1d4ed8'}
                                      onChange={(e) => actualizarBloqueBoton(bloque.id, 'colorBoton', e.target.value)}
                                      className="w-10 h-10 rounded cursor-pointer border"
                                    />
                                    <div>
                                      <span className="block text-xs font-bold text-slate-700">Color Principal</span>
                                      <span className="text-[10px] text-slate-400">{bloque.colorBoton}</span>
                                    </div>
                                  </div>

                                  <div className="flex items-center gap-3 bg-white p-3 rounded-xl border border-slate-200">
                                    <input
                                      type="color"
                                      value={bloque.colorTexto || '#ffffff'}
                                      onChange={(e) => actualizarBloqueBoton(bloque.id, 'colorTexto', e.target.value)}
                                      className="w-10 h-10 rounded cursor-pointer border"
                                    />
                                    <div>
                                      <span className="block text-xs font-bold text-slate-700">Color de Texto</span>
                                      <span className="text-[10px] text-slate-400">{bloque.colorTexto}</span>
                                    </div>
                                  </div>

                                  {bloque.estilo === '4d' && (
                                    <div className="flex items-center gap-3 bg-white p-3 rounded-xl border border-blue-200">
                                      <input
                                        type="color"
                                        value={bloque.colorIluminacion || '#60a5fa'}
                                        onChange={(e) => actualizarBloqueBoton(bloque.id, 'colorIluminacion', e.target.value)}
                                        className="w-10 h-10 rounded cursor-pointer border"
                                      />
                                      <div>
                                        <span className="block text-xs font-bold text-blue-700">Luz 4D (Resplandor)</span>
                                        <span className="text-[10px] text-slate-400">{bloque.colorIluminacion || '#60a5fa'}</span>
                                      </div>
                                    </div>
                                  )}
                                </div>
                              </div>

                              <div className={`p-6 bg-slate-100/70 rounded-xl border border-dashed border-slate-300 flex ${bloque.alineacion === 'centro' ? 'justify-center' : bloque.alineacion === 'derecha' ? 'justify-end' : 'justify-start'}`}>
                                <a
                                  href={bloque.url || '#'}
                                  target="_blank"
                                  rel="noreferrer"
                                  onClick={(e) => e.preventDefault()}
                                  style={{
                                    backgroundColor: bloque.estilo === 'contorno' ? 'transparent' : bloque.colorBoton,
                                    color: bloque.estilo === 'contorno' ? bloque.colorBoton : bloque.colorTexto,
                                    borderColor: bloque.colorBoton,
                                    boxShadow: bloque.estilo === '3d' 
                                      ? `0 6px 0 0 ${bloque.colorBoton}aa, 0 10px 15px -3px rgba(0,0,0,0.2)` 
                                      : bloque.estilo === '4d'
                                      ? `0 0 25px 4px ${bloque.colorIluminacion || bloque.colorBoton}, inset 0 0 12px rgba(255,255,255,0.4)`
                                      : undefined
                                  }}
                                  className={`px-8 py-4 font-sans font-black text-sm uppercase tracking-wider transition-all inline-block ${
                                    bloque.estilo === 'pildora' ? 'rounded-full shadow-md' :
                                    bloque.estilo === 'contorno' ? 'rounded-xl border-2' :
                                    bloque.estilo === '3d' ? 'rounded-xl -translate-y-1' :
                                    bloque.estilo === '4d' ? 'rounded-full border border-white/40' :
                                    'rounded-xl shadow-md'
                                  }`}
                                >
                                  {bloque.texto || 'Botón'}
                                </a>
                              </div>
                            </div>
                          )}

                          {bloque.tipo === 'imagen' && (
                            <div className="p-6 rounded-b-2xl" style={{ backgroundColor: bloque.colorFondo || '#ffffff' }}>
                              <div className="flex flex-col gap-6 mb-6">
                                <div className="bg-slate-50 p-6 rounded-xl border border-slate-200 flex flex-col md:flex-row gap-6 items-start md:items-center justify-between">
                                  <div className="w-full md:w-1/2">
                                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-3">Opción 1: Subir Archivo Local</label>
                                    <div className="relative group">
                                      <input type="file" accept="image/*" onChange={(e) => manejarSubidaImagen(e, bloque.id)} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" disabled={subiendoImagen[bloque.id]} />
                                      <button type="button" className={`w-full px-6 py-4 rounded-xl text-sm font-bold shadow-sm flex justify-center items-center gap-3 transition-colors ${subiendoImagen[bloque.id] ? 'bg-slate-200 text-slate-500 cursor-not-allowed' : 'bg-blue-600 text-white hover:bg-blue-700 group-hover:shadow-md'}`}>
                                        {subiendoImagen[bloque.id] ? <RefreshCcw size={20} className="animate-spin" /> : <Upload size={20} />}
                                        {subiendoImagen[bloque.id] ? 'Subiendo archivo a la nube...' : 'Seleccionar Imagen de tu PC'}
                                      </button>
                                    </div>
                                  </div>
                                  <div className="hidden md:flex flex-col items-center justify-center px-2">
                                    <div className="h-6 w-px bg-slate-300"></div>
                                    <span className="text-[10px] font-black text-slate-400 uppercase tracking-widest my-2">O</span>
                                    <div className="h-6 w-px bg-slate-300"></div>
                                  </div>
                                  <div className="w-full md:w-1/2">
                                    <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-3">Opción 2: Enlace Externo (URL)</label>
                                    <input type="text" value={bloque.url} onChange={(e) => { const nuevasPaginas = [...paginas]; const b = nuevasPaginas[indicePaginaActual].bloques.find(bl => bl.id === bloque.id) as BloqueImagen; if (b) { b.url = e.target.value; autoGuardar(nuevasPaginas); } }} className="w-full p-3.5 bg-white border border-slate-200 rounded-xl text-sm font-semibold outline-none focus:border-blue-500 transition-colors" placeholder="https://ejemplo.com/foto.jpg" />
                                  </div>
                                </div>

                                <label className="flex items-center gap-3 text-sm font-medium text-slate-700 cursor-pointer p-3 hover:bg-slate-50 rounded-xl bg-white border border-slate-200 w-fit shadow-sm">
                                  <input type="checkbox" checked={bloque.pantallaCompleta} onChange={(e) => { const nuevasPaginas = [...paginas]; const b = nuevasPaginas[indicePaginaActual].bloques.find(bl => bl.id === bloque.id) as BloqueImagen; if (b) { b.pantallaCompleta = e.target.checked; autoGuardar(nuevasPaginas); } }} className="w-5 h-5 rounded border-slate-300 text-blue-600 focus:ring-blue-500" />
                                  <span className="font-bold">Imagen de Sangrado Completo (Full Bleed)</span>
                                </label>
                              </div>

                              {bloque.url && (
                                <div className="mt-6 rounded-2xl overflow-hidden border-4 border-slate-100 bg-slate-50 flex justify-center items-center shadow-inner relative">
                                  <div className="absolute top-4 left-4 bg-white/90 backdrop-blur px-3 py-1.5 rounded-lg border border-slate-200 shadow-sm text-xs font-bold uppercase tracking-widest text-slate-600">Vista Previa</div>
                                  <img src={bloque.url} alt="Vista previa" className="w-full h-auto object-contain max-h-100" />
                                </div>
                              )}
                            </div>
                          )}

                          {/* TABLA CON VISTA PREVIA ULTRA COMPACTA REAL */}
                          {bloque.tipo === 'tabla' && (
                            <div className="p-6 rounded-b-2xl" style={{ backgroundColor: bloque.colorFondo || '#ffffff' }}>
                              <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-4">
                                <input 
                                  type="text" 
                                  value={bloque.nombre} 
                                  onChange={(e) => {
                                    const nuevasPaginas = [...paginas];
                                    const b = nuevasPaginas[indicePaginaActual].bloques.find(bl => bl.id === bloque.id) as BloqueTabla;
                                    if (b) { b.nombre = e.target.value; autoGuardar(nuevasPaginas); }
                                  }}
                                  className="text-lg font-bold text-slate-800 bg-transparent outline-none border-b-2 border-transparent focus:border-blue-500 w-full md:w-2/3"
                                  placeholder="Título de la tabla (Ej. Datos Ing. Civil)"
                                />
                                
                                <button 
                                  onClick={() => alternarVisibilidadTabla(bloque.id)}
                                  className={`flex items-center gap-2 px-4 py-2 rounded-lg text-xs font-bold transition-all shadow-sm cursor-pointer ${bloque.oculto ? 'bg-amber-100 text-amber-700 hover:bg-amber-200' : 'bg-slate-100 text-slate-600 hover:bg-slate-200'}`}
                                >
                                  {bloque.oculto ? <EyeOff size={16} /> : <Eye size={16} />}
                                  {bloque.oculto ? 'Tabla Oculta' : 'Tabla Visible'}
                                </button>
                              </div>
                              
                              <div 
                                className="overflow-x-auto border border-slate-200 rounded-xl mx-auto transition-all"
                                style={{ 
                                  backgroundColor: bloque.colorFondo || '#ffffff',
                                  width: `${Math.min(100, bloque.escala || 100)}%`,
                                  zoom: (bloque.escala || 100) > 100 ? (bloque.escala || 100) / 100 : 1
                                }}
                              >
                                <table 
                                  className="w-full text-left border-collapse" 
                                  style={{ 
                                    backgroundColor: bloque.colorFondo || '#ffffff',
                                    fontSize: `${bloque.tamanoTextoTabla ?? 14}px`
                                  }}
                                >
                                  <thead>
                                    <tr className="bg-slate-100/80 border-b border-slate-200">
                                      {bloque.data[0]?.map((_, cIndex) => {
                                        const anchoCol = bloque.anchosColumnas?.[cIndex];
                                        return (
                                          <th 
                                            key={cIndex} 
                                            className="p-1 text-center border-r border-slate-200 last:border-r-0"
                                            style={anchoCol ? { width: `${anchoCol}px`, minWidth: `${anchoCol}px`, maxWidth: `${anchoCol}px` } : undefined}
                                          >
                                            <div className="flex flex-col items-center gap-1">
                                              <div className="flex items-center justify-center gap-0.5 flex-nowrap whitespace-nowrap">
                                                <button
                                                  type="button"
                                                  onClick={() => moverColumnaTabla(bloque.id, cIndex, 'izq')}
                                                  disabled={cIndex === 0}
                                                  className="p-0.5 text-slate-500 hover:text-blue-600 disabled:opacity-25 cursor-pointer"
                                                  title="Mover columna a la izquierda"
                                                >
                                                  <MoveLeft size={11} />
                                                </button>
                                                <button
                                                  type="button"
                                                  onClick={() => moverColumnaTabla(bloque.id, cIndex, 'der')}
                                                  disabled={cIndex === (bloque.data[0]?.length || 1) - 1}
                                                  className="p-0.5 text-slate-500 hover:text-blue-600 disabled:opacity-25 cursor-pointer"
                                                  title="Mover columna a la derecha"
                                                >
                                                  <MoveRight size={11} />
                                                </button>
                                                <input
                                                  type="color"
                                                  value={bloque.coloresColumnas?.[cIndex] || bloque.colorFondo || '#ffffff'}
                                                  onChange={(e) => actualizarColorColumnaTabla(bloque.id, cIndex, e.target.value)}
                                                  className="w-4 h-4 rounded cursor-pointer border-0 p-0 shrink-0"
                                                  title="Color de esta columna"
                                                />
                                                <button
                                                  type="button"
                                                  onClick={() => insertarColumnaEnPosicion(bloque.id, cIndex)}
                                                  className="text-[10px] font-bold text-blue-600 hover:bg-blue-50 px-1 py-0.5 rounded cursor-pointer"
                                                  title="Insertar columna a la derecha de esta"
                                                >
                                                  +Col
                                                </button>
                                                <button
                                                  type="button"
                                                  onClick={() => eliminarColumnaTabla(bloque.id, cIndex)}
                                                  className="text-[10px] font-bold text-red-500 hover:bg-red-50 px-1 py-0.5 rounded inline-flex items-center cursor-pointer"
                                                  title="Eliminar esta columna"
                                                >
                                                  <X size={11} />
                                                </button>
                                              </div>
                                              <div className="flex items-center gap-1 bg-white px-1.5 py-0.5 rounded border border-slate-200 text-[10px] text-slate-500">
                                                <span>Ancho:</span>
                                                <input
                                                  type="number"
                                                  min="0"
                                                  max="800"
                                                  placeholder="Auto"
                                                  value={bloque.anchosColumnas?.[cIndex] || ''}
                                                  onChange={(e) => actualizarAnchoColumnaTabla(bloque.id, cIndex, Number(e.target.value))}
                                                  className="w-11 text-center font-bold text-slate-700 outline-none bg-transparent"
                                                  title="Ancho en píxeles para esta columna (0 o vacío = Automático)"
                                                />
                                                <span>px</span>
                                              </div>
                                            </div>
                                          </th>
                                        );
                                      })}
                                      <th className="w-44 bg-slate-100/80 text-center text-[10px] font-bold uppercase text-slate-400 px-2">Filas / Alto</th>
                                    </tr>
                                  </thead>
                                  <tbody className="divide-y divide-slate-200/60">
                                    {bloque.data.map((fila, fIndex) => {
                                      const colorFila = bloque.coloresFilas?.[fIndex];
                                      const padGeneral = bloque.paddingFila ?? 6;
                                      const altoIndividual = bloque.altosFilas?.[fIndex];
                                      const padVertical = altoIndividual !== undefined && altoIndividual > 0
                                        ? Math.max(0, Math.round((altoIndividual - (bloque.tamanoTextoTabla ?? 14)) / 2))
                                        : padGeneral;
                                      const fontSizeCelda = fIndex === 0
                                        ? Math.max(9, (bloque.tamanoTextoTabla ?? 14) - 1)
                                        : (bloque.tamanoTextoTabla ?? 14);
                                      const alturaExactaInput = altoIndividual !== undefined && altoIndividual > 0
                                        ? altoIndividual
                                        : Math.max(16, Math.round(fontSizeCelda * 1.15) + padVertical * 2);

                                      return (
                                        <tr 
                                          key={fIndex} 
                                          className={fIndex === 0 ? "font-bold" : ""}
                                          style={{ 
                                            backgroundColor: colorFila || bloque.colorFondo || '#ffffff',
                                            height: `${alturaExactaInput}px`
                                          }}
                                        >
                                          {fila.map((celda, cIndex) => {
                                            const colorCol = bloque.coloresColumnas?.[cIndex];
                                            const bgCelda = colorFila || colorCol || bloque.colorFondo || '#ffffff';
                                            const anchoCol = bloque.anchosColumnas?.[cIndex];
                                            return (
                                              <td 
                                                key={cIndex} 
                                                className="border-r border-slate-200/60 p-0 last:border-r-0 align-middle"
                                                style={{ 
                                                  backgroundColor: bgCelda,
                                                  height: `${alturaExactaInput}px`,
                                                  ...(anchoCol ? { width: `${anchoCol}px`, minWidth: `${anchoCol}px`, maxWidth: `${anchoCol}px` } : {})
                                                }}
                                              >
                                                <input 
                                                  type="text" 
                                                  value={celda} 
                                                  onChange={(e) => actualizarCeldaTabla(bloque.id, fIndex, cIndex, e.target.value)}
                                                  style={{ 
                                                    backgroundColor: 'transparent',
                                                    fontSize: `${fontSizeCelda}px`,
                                                    lineHeight: 1.1,
                                                    height: `${alturaExactaInput}px`,
                                                    minHeight: '0px',
                                                    paddingTop: `${padVertical}px`,
                                                    paddingBottom: `${padVertical}px`,
                                                    paddingLeft: '14px',
                                                    paddingRight: '14px',
                                                    boxSizing: 'border-box'
                                                  }}
                                                  className={`w-full block m-0 outline-none focus:ring-2 focus:ring-inset focus:ring-blue-500 transition-all ${fIndex === 0 ? 'font-bold uppercase tracking-wider text-slate-800' : (cIndex === 0 ? 'font-semibold text-slate-800' : 'text-slate-700')}`}
                                                />
                                              </td>
                                            );
                                          })}
                                          <td 
                                            className="py-0 px-1.5 text-center w-44 bg-slate-50/90 border-l border-slate-200 align-middle"
                                            style={{ height: `${alturaExactaInput}px` }}
                                          >
                                            <div className="flex items-center justify-center gap-1 flex-nowrap whitespace-nowrap leading-none">
                                              <button
                                                type="button"
                                                onClick={() => moverFilaTabla(bloque.id, fIndex, 'arriba')}
                                                disabled={fIndex === 0}
                                                className="p-0.5 text-slate-500 hover:text-blue-600 disabled:opacity-25 cursor-pointer"
                                                title="Subir fila"
                                              >
                                                <ArrowUp size={11} />
                                              </button>
                                              <button
                                                type="button"
                                                onClick={() => moverFilaTabla(bloque.id, fIndex, 'abajo')}
                                                disabled={fIndex === bloque.data.length - 1}
                                                className="p-0.5 text-slate-500 hover:text-blue-600 disabled:opacity-25 cursor-pointer"
                                                title="Bajar fila"
                                              >
                                                <ArrowDown size={11} />
                                              </button>
                                              <input
                                                type="color"
                                                value={bloque.coloresFilas?.[fIndex] || bloque.colorFondo || '#ffffff'}
                                                onChange={(e) => actualizarColorFilaTabla(bloque.id, fIndex, e.target.value)}
                                                className="w-3.5 h-3.5 rounded cursor-pointer border-0 p-0 shrink-0"
                                                title="Color de esta fila"
                                              />
                                              <input
                                                type="number"
                                                min="0"
                                                max="150"
                                                placeholder={`${alturaExactaInput}`}
                                                value={bloque.altosFilas?.[fIndex] || ''}
                                                onChange={(e) => actualizarAltoFilaTabla(bloque.id, fIndex, Number(e.target.value))}
                                                className="w-9 h-4 text-[10px] font-bold text-center bg-white border border-slate-200 rounded p-0 outline-none leading-none"
                                                title="Alto individual de esta fila en px (vacío = usa el espacio general)"
                                              />
                                              <button
                                                type="button"
                                                onClick={() => insertarFilaEnPosicion(bloque.id, fIndex)}
                                                className="text-[10px] font-bold text-blue-600 hover:bg-blue-50 px-1 py-0 rounded leading-none cursor-pointer"
                                                title="Insertar fila debajo de esta"
                                              >
                                                +F
                                              </button>
                                              {fIndex > 0 && (
                                                <button
                                                  type="button"
                                                  onClick={() => eliminarFilaTabla(bloque.id, fIndex)}
                                                  className="p-0.5 text-red-400 hover:text-red-600 hover:bg-red-50 rounded transition-colors cursor-pointer"
                                                  title="Eliminar esta fila"
                                                >
                                                  <Trash2 size={11} />
                                                </button>
                                              )}
                                            </div>
                                          </td>
                                        </tr>
                                      );
                                    })}
                                  </tbody>
                                </table>
                              </div>
                              <div className="flex gap-2 mt-4">
                                <button onClick={() => agregarFilaTabla(bloque.id)} className="text-xs font-bold bg-slate-100 hover:bg-slate-200 text-slate-600 px-4 py-2 rounded-lg transition-colors flex items-center gap-2 cursor-pointer"><Plus size={14} /> Fila al Final</button>
                                <button onClick={() => agregarColumnaTabla(bloque.id)} className="text-xs font-bold bg-slate-100 hover:bg-slate-200 text-slate-600 px-4 py-2 rounded-lg transition-colors flex items-center gap-2 cursor-pointer"><Plus size={14} /> Columna al Final</button>
                              </div>

                              <details className="mt-5 pt-4 border-t border-slate-100 group" open={!!bloque.pieBloque}>
                                <summary className="text-xs font-bold uppercase tracking-widest text-slate-500 hover:text-blue-600 cursor-pointer select-none flex items-center gap-2">
                                  <Type size={14} className="text-blue-600" />
                                  <span>+ Agregar texto / fuente / nota justo debajo de esta tabla (sin espacio de bloque)</span>
                                </summary>
                                <div className="mt-3 bg-white border border-slate-200 rounded-xl [&_.ql-toolbar]:border-none [&_.ql-toolbar]:bg-slate-50 [&_.ql-toolbar]:rounded-t-xl [&_.ql-container]:border-none [&_.ql-editor]:min-h-16 [&_.ql-editor]:text-sm">
                                  <ReactQuill
                                    theme="snow"
                                    modules={quillModules}
                                    value={bloque.pieBloque || ''}
                                    onChange={(val) => actualizarPropiedadTabla(bloque.id, 'pieBloque', val)}
                                    placeholder="Ej: Fuente: Elaboración propia ACOFI 2025..."
                                  />
                                </div>
                              </details>
                            </div>
                          )}

                          {bloque.tipo === 'grafica' && (
                            <div className="p-6 rounded-b-2xl" style={{ backgroundColor: bloque.colorFondo || '#ffffff' }}>
                              <input 
                                type="text" 
                                value={bloque.nombre} 
                                onChange={(e) => actualizarGrafica(bloque.id, 'nombre', e.target.value)}
                                className="text-lg font-bold text-slate-800 mb-6 bg-transparent outline-none border-b-2 border-transparent focus:border-blue-500 w-full"
                                placeholder="Título de la gráfica o mapa"
                              />
                              
                              <div className="grid grid-cols-1 md:grid-cols-2 gap-6 bg-slate-50/50 p-6 rounded-xl border border-slate-200">
                                <div>
                                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Fuente de Datos</label>
                                  <select 
                                    value={bloque.tablaRef} 
                                    onChange={(e) => actualizarGrafica(bloque.id, 'tablaRef', e.target.value)}
                                    className="w-full p-3 bg-white border border-slate-200 rounded-xl text-sm font-semibold outline-none focus:border-blue-500"
                                  >
                                    <option value="">-- Seleccionar Tabla --</option>
                                    {todasLasTablasAdmin.map(t => (
                                      <option key={t.id} value={t.id}>{t.nombre || 'Tabla sin nombre'} ({t.nombrePagina}) {t.oculto ? '👁️‍🗨️' : ''}</option>
                                    ))}
                                  </select>
                                </div>
                                
                                <div>
                                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Tipo de Visualización Premium</label>
                                  <select 
                                    value={bloque.tipoGrafica} 
                                    onChange={(e) => actualizarGrafica(bloque.id, 'tipoGrafica', e.target.value as TipoGrafica)}
                                    className="w-full p-3 bg-white border border-slate-200 rounded-xl text-sm font-semibold outline-none focus:border-blue-500"
                                  >
                                    <optgroup label="Gráficas Circulares 3D y Rellenas (Sin Hueco / Con Hueco)">
                                      <option value="torta_3d_rellena">Torta 3D Rellena 100% (Sin hueco en el centro)</option>
                                      <option value="torta_rellena">Torta Plana Rellena 100% (Sin hueco central)</option>
                                      <option value="rosa_polar">Rosa Polar (Triángulos desde el eje central sin hueco)</option>
                                      <option value="torta">Gráfica de Torta Clásica</option>
                                      <option value="dona">Gráfica de Dona Moderna</option>
                                      <option value="torta_4d">Dona 3D / 4D (Relieve & Glow)</option>
                                      <option value="semicirculo">Semicírculo / Tacómetro (180°)</option>
                                      <option value="barras_radiales">Anillos Radiales Concéntricos</option>
                                    </optgroup>
                                    <optgroup label="Gráficas de Barras">
                                      <option value="barras">Barras Verticales Clásicas</option>
                                      <option value="barras_horizontales">Barras Horizontales (Ranking)</option>
                                      <option value="barras_apiladas">Barras Apiladas Comparativas</option>
                                      <option value="barras_3d_apiladas">Barras 3D Apiladas con Sombra</option>
                                      <option value="barras_4d">Barras 3D / 4D (Cilindro Iluminado)</option>
                                    </optgroup>
                                    <optgroup label="Tendencias, Jerarquía y Mapas">
                                      <option value="lineas">Gráfica de Líneas Rectas</option>
                                      <option value="lineas_curvas">Gráfica de Líneas Curvas Suaves (Spline)</option>
                                      <option value="area">Gráfica de Área</option>
                                      <option value="area_apilada">Gráfica de Áreas Apiladas</option>
                                      <option value="combinada">Gráfica Combinada (Barras + Línea)</option>
                                      <option value="dispersion">Gráfica de Dispersión (Puntos)</option>
                                      <option value="radar">Gráfica de Radar (Araña)</option>
                                      <option value="radar_relleno">Radar Circular Relleno (Triángulos en Red)</option>
                                      <option value="embudo">Gráfica de Embudo (Funnel)</option>
                                      <option value="treemap">Mapa de Árbol Jerárquico (Treemap)</option>
                                      <option value="mapa_colombia">Mapa Geográfico de Colombia</option>
                                    </optgroup>
                                  </select>
                                </div>

                                {bloque.tablaRef && (() => {
                                  const { datos, cabeceras, ejesYUsar } = formatearDatosParaGrafica(bloque);
                                  if (cabeceras.length === 0) {
                                    return <div className="col-span-2 text-sm text-red-500">Tabla fuente no válida o vacía.</div>;
                                  }
                                  
                                  const usaColorPorFila = esGraficaPorSegmento(bloque.tipoGrafica) || (['barras', 'barras_horizontales', 'barras_4d'].includes(bloque.tipoGrafica) && ejesYUsar.length === 1);

                                  return (
                                    <>
                                      <div className="col-span-1 md:col-span-2 border-t border-slate-200 my-2"></div>
                                      <div>
                                        <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">
                                          {bloque.tipoGrafica === 'mapa_colombia' ? 'Columna de Ubicación (Departamento/Ciudad)' : 'Eje X (Categorías / Periodos)'}
                                        </label>
                                        <select 
                                          value={bloque.ejeX} 
                                          onChange={(e) => actualizarGrafica(bloque.id, 'ejeX', Number(e.target.value))}
                                          className="w-full p-3 bg-white border border-slate-200 rounded-xl text-sm font-semibold outline-none focus:border-blue-500"
                                        >
                                          {cabeceras.map((cab, i) => (
                                            <option key={i} value={i}>{cab || `Columna ${i + 1}`}</option>
                                          ))}
                                        </select>
                                      </div>
                                      
                                      <div>
                                        <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">
                                          {bloque.tipoGrafica === 'mapa_colombia' ? 'Columna de Valores a Mapear' : 'Ejes Y (Valores Numéricos o %)'}
                                        </label>
                                        <div className="space-y-2 max-h-32 overflow-y-auto p-3 bg-white border border-slate-200 rounded-xl">
                                          {cabeceras.map((cab, i) => {
                                            if (i === Number(bloque.ejeX)) return null; 
                                            const isChecked = (bloque.ejesY || []).includes(i);
                                            return (
                                              <label key={i} className="flex items-center gap-3 text-sm font-medium text-slate-700 cursor-pointer hover:bg-slate-50 p-1.5 rounded-lg transition-colors">
                                                <input 
                                                  type="checkbox" 
                                                  checked={isChecked}
                                                  onChange={(e) => {
                                                    let nuevosEjesY = [...(bloque.ejesY || [])];
                                                    if (e.target.checked) {
                                                      if (bloque.tipoGrafica === 'mapa_colombia') nuevosEjesY = [i];
                                                      else nuevosEjesY.push(i);
                                                    } else {
                                                      nuevosEjesY = nuevosEjesY.filter(val => val !== i);
                                                    }
                                                    actualizarGrafica(bloque.id, 'ejesY', nuevosEjesY);
                                                  }}
                                                  className="w-4 h-4 rounded border-slate-300 text-blue-600 focus:ring-blue-500"
                                                />
                                                {cab || `Columna ${i + 1}`}
                                              </label>
                                            );
                                          })}
                                          {(!bloque.ejesY || bloque.ejesY.length === 0) && (
                                            <p className="text-xs text-slate-400 italic">⚠️ El sistema usará la primera columna numérica por defecto.</p>
                                          )}
                                        </div>
                                      </div>

                                      <div className="col-span-1 md:col-span-2 pt-2">
                                        <label className="flex items-center gap-3 text-sm font-bold text-slate-800 cursor-pointer bg-white p-4 rounded-xl border border-slate-200 hover:border-blue-400 transition-colors shadow-sm">
                                          <input
                                            type="checkbox"
                                            checked={!!bloque.mostrarValores}
                                            onChange={(e) => actualizarGrafica(bloque.id, 'mostrarValores', e.target.checked)}
                                            className="w-5 h-5 rounded text-blue-600 focus:ring-blue-500"
                                          />
                                          {bloque.tipoGrafica === 'mapa_colombia'
                                            ? 'Mostrar tabla lateral de valores por departamento (Desmárcala para dejar solo el mapa limpio con valores al pasar el cursor)'
                                            : 'Mostrar valores numéricos / estadísticas fijos encima de la gráfica (Sin necesidad de pasar el mouse)'}
                                        </label>
                                      </div>

                                      <div className="col-span-1 md:col-span-2 mt-2 pt-6 border-t border-slate-200">
                                        <h4 className="text-xs font-bold text-slate-500 uppercase tracking-widest mb-2 flex items-center gap-2">
                                          <Palette size={16} /> Personalización Exacta de Colores ({bloque.tipoGrafica === 'mapa_colombia' ? 1 : (usaColorPorFila ? datos.length : ejesYUsar.length)} elementos activos)
                                        </h4>
                                        <p className="text-xs text-slate-400 mb-4">
                                          💡 También puedes hacer <strong>clic directamente sobre cualquier barra o segmento en la Vista Previa</strong> de abajo para cambiar su color.
                                        </p>
                                        
                                        {bloque.tipoGrafica === 'mapa_colombia' ? (
                                          <div className="flex items-center gap-4 bg-white p-4 rounded-xl border border-slate-200 w-fit">
                                            <input 
                                              type="color" 
                                              value={bloque.coloresCustom?.[0] || edicionActiva.color_principal || '#93c5fd'} 
                                              onChange={(e) => actualizarGraficaColor(bloque.id, 0, e.target.value)} 
                                              className="w-10 h-10 cursor-pointer rounded-lg border-2"
                                            />
                                            <div>
                                              <span className="block text-sm font-bold text-slate-800">Color principal del Mapa</span>
                                              <span className="block text-xs text-slate-500">Tono fuerte del gradiente</span>
                                            </div>
                                          </div>
                                        ) : usaColorPorFila ? (
                                          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-3 max-h-60 overflow-y-auto p-2 bg-white rounded-xl border border-slate-200">
                                            {datos.map((d: any, idx: number) => (
                                              <div key={idx} className="flex items-center gap-2.5 p-2 rounded-lg border border-slate-100 hover:border-blue-300 transition-colors">
                                                <input 
                                                  type="color" 
                                                  value={bloque.coloresCustom?.[idx] || COLORES_GRAFICAS_DEFAULT[idx % COLORES_GRAFICAS_DEFAULT.length]} 
                                                  onChange={(e) => actualizarGraficaColor(bloque.id, idx, e.target.value)} 
                                                  className="w-7 h-7 cursor-pointer rounded-md border shrink-0"
                                                  title={`Cambiar color de ${d.name}`}
                                                />
                                                <span className="text-xs font-bold text-slate-700 truncate" title={d.name}>{d.name}</span>
                                              </div>
                                            ))}
                                          </div>
                                        ) : (
                                          <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                                            {ejesYUsar.map((ejeYIdx, idx) => {
                                               const nombreEje = cabeceras[ejeYIdx] || `Serie ${idx + 1}`;
                                               return (
                                                  <div key={idx} className="flex items-center gap-3 bg-white p-3 rounded-xl border border-slate-200">
                                                    <input 
                                                      type="color" 
                                                      value={bloque.coloresCustom?.[idx] || COLORES_GRAFICAS_DEFAULT[idx % COLORES_GRAFICAS_DEFAULT.length]} 
                                                      onChange={(e) => actualizarGraficaColor(bloque.id, idx, e.target.value)} 
                                                      className="w-8 h-8 cursor-pointer rounded-md border shrink-0"
                                                    />
                                                    <span className="text-sm font-bold text-slate-700 truncate">{nombreEje}</span>
                                                  </div>
                                               );
                                            })}
                                          </div>
                                        )}
                                      </div>
                                    </>
                                  );
                                })()}
                              </div>

                              {bloque.tablaRef && (
                                <div 
                                  className="mt-8 p-6 border-2 border-dashed border-slate-200 rounded-xl relative mx-auto transition-all"
                                  style={{ 
                                    backgroundColor: bloque.colorFondo || '#ffffff',
                                    width: `${Math.min(100, bloque.escala || 100)}%`,
                                    zoom: (bloque.escala || 100) > 100 ? (bloque.escala || 100) / 100 : 1
                                  }}
                                >
                                  <div className="flex flex-wrap items-center justify-between gap-2 mb-4">
                                    <h4 className="text-xs font-bold uppercase text-slate-400 tracking-widest">
                                      Vista Previa Interactiva ({bloque.tipoGrafica}) — Haz clic en un elemento de la gráfica para cambiar su color
                                    </h4>
                                    {selectorColorActivo && selectorColorActivo.bloqueId === bloque.id && (
                                      <div className="flex items-center gap-3 bg-slate-900 text-white px-4 py-2 rounded-xl shadow-lg animate-in fade-in duration-200">
                                        <span className="text-xs font-bold">Color de &quot;{selectorColorActivo.etiqueta}&quot;:</span>
                                        <input
                                          type="color"
                                          value={bloque.coloresCustom?.[selectorColorActivo.indice] || COLORES_GRAFICAS_DEFAULT[selectorColorActivo.indice % COLORES_GRAFICAS_DEFAULT.length]}
                                          onChange={(e) => actualizarGraficaColor(bloque.id, selectorColorActivo.indice, e.target.value)}
                                          className="w-7 h-7 rounded cursor-pointer border-0"
                                        />
                                        <button
                                          type="button"
                                          onClick={() => setSelectorColorActivo(null)}
                                          className="text-slate-400 hover:text-white cursor-pointer"
                                        >
                                          <X size={14} />
                                        </button>
                                      </div>
                                    )}
                                  </div>
                                  {renderizarGrafica(bloque)}
                                  {bloque.pieBloque && (
                                    <div 
                                      className="mt-3 pt-2 border-t border-slate-100 prose prose-sm max-w-none text-slate-600"
                                      dangerouslySetInnerHTML={{ __html: renderizarBotonesEnTexto(bloque.pieBloque) }}
                                    />
                                  )}
                                </div>
                              )}

                              <details className="mt-5 pt-4 border-t border-slate-100 group" open={!!bloque.pieBloque}>
                                <summary className="text-xs font-bold uppercase tracking-widest text-slate-500 hover:text-blue-600 cursor-pointer select-none flex items-center gap-2">
                                  <Type size={14} className="text-blue-600" />
                                  <span>+ Agregar texto / fuente / nota justo debajo de esta gráfica (sin espacio de bloque)</span>
                                </summary>
                                <div className="mt-3 bg-white border border-slate-200 rounded-xl [&_.ql-toolbar]:border-none [&_.ql-toolbar]:bg-slate-50 [&_.ql-toolbar]:rounded-t-xl [&_.ql-container]:border-none [&_.ql-editor]:min-h-16 [&_.ql-editor]:text-sm">
                                  <ReactQuill
                                    theme="snow"
                                    modules={quillModules}
                                    value={bloque.pieBloque || ''}
                                    onChange={(val) => actualizarGrafica(bloque.id, 'pieBloque', val)}
                                    placeholder="Ej: Nota: Cifras expresadas en porcentaje nacional..."
                                  />
                                </div>
                              </details>

                            </div>
                          )}
                        </div>
                      </div>
                    ))}

                    <div className="flex flex-wrap items-center justify-center gap-4 py-8 border-t-2 border-dashed border-slate-200 mt-12">
                      <span className="text-xs font-bold uppercase tracking-widest text-slate-400 mr-2 w-full text-center md:w-auto">Insertar Bloque:</span>
                      <button onClick={agregarBloqueTexto} className="flex items-center gap-2 bg-white border-2 border-slate-200 text-slate-600 hover:border-blue-600 hover:text-blue-600 px-6 py-3 rounded-xl text-sm font-bold shadow-sm transition-all hover:-translate-y-1 cursor-pointer"><Type size={18} /> Texto (1, 2 o 3 Col)</button>
                      <button onClick={agregarBloqueFirma} className="flex items-center gap-2 bg-white border-2 border-slate-200 text-slate-600 hover:border-amber-600 hover:text-amber-600 px-6 py-3 rounded-xl text-sm font-bold shadow-sm transition-all hover:-translate-y-1 cursor-pointer"><PenTool size={18} /> Firma(s)</button>
                      <button onClick={agregarBloqueBoton} className="flex items-center gap-2 bg-white border-2 border-slate-200 text-slate-600 hover:border-indigo-600 hover:text-indigo-600 px-6 py-3 rounded-xl text-sm font-bold shadow-sm transition-all hover:-translate-y-1 cursor-pointer"><MousePointerClick size={18} /> Botón Enlace</button>
                      <button onClick={agregarBloqueImagen} className="flex items-center gap-2 bg-white border-2 border-slate-200 text-slate-600 hover:border-pink-600 hover:text-pink-600 px-6 py-3 rounded-xl text-sm font-bold shadow-sm transition-all hover:-translate-y-1 cursor-pointer"><ImageIcon size={18} /> Imagen</button>
                      <button onClick={agregarBloqueTabla} className="flex items-center gap-2 bg-white border-2 border-slate-200 text-slate-600 hover:border-emerald-600 hover:text-emerald-600 px-6 py-3 rounded-xl text-sm font-bold shadow-sm transition-all hover:-translate-y-1 cursor-pointer"><TableIcon size={18} /> Tabla</button>
                      <button onClick={agregarBloqueGrafica} className="flex items-center gap-2 bg-white border-2 border-slate-200 text-slate-600 hover:border-purple-600 hover:text-purple-600 px-6 py-3 rounded-xl text-sm font-bold shadow-sm transition-all hover:-translate-y-1 cursor-pointer"><BarChart3 size={18} /> Gráfica o Mapa</button>
                    </div>
                  </div>

                </div>
              )}

            </div>
          ) : (
            <div className="flex h-full items-center justify-center"><h3 className="text-2xl font-black text-slate-800">Ninguna edición activa</h3></div>
          )}
        </main>
      )}

      {/* MODAL PARA INYECTAR LOTES DE PÁGINAS EN FORMATO JSON */}
      {modalImportarJsonAbierto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 no-print">
          <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm" onClick={() => setModalImportarJsonAbierto(false)}></div>
          <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-2xl relative z-10 animate-in zoom-in-95 duration-200 overflow-hidden">
            <div className="h-2 w-full bg-emerald-500"></div>
            <div className="p-8">
              <div className="flex justify-between items-start mb-4">
                <div>
                  <h3 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
                    <Database size={22} className="text-emerald-600" /> Inyector de Lotes de Páginas (JSON)
                  </h3>
                  <p className="text-xs text-slate-500 mt-1">
                    Pega aquí el arreglo JSON con las páginas, textos, tablas y gráficas generadas para insertarlas automáticamente.
                  </p>
                </div>
                <button onClick={() => setModalImportarJsonAbierto(false)} className="text-slate-400 hover:text-slate-600 p-1 cursor-pointer"><X size={20} /></button>
              </div>

              <div className="space-y-4 mt-4">
                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">¿En dónde deseas insertar este lote?</label>
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                    <button
                      type="button"
                      onClick={() => setModoInsercionJson('final')}
                      className={`p-3 rounded-xl border text-xs font-bold text-left transition-all cursor-pointer ${modoInsercionJson === 'final' ? 'border-emerald-600 bg-emerald-50 text-emerald-900' : 'border-slate-200 text-slate-600 hover:bg-slate-50'}`}
                    >
                      ➕ Al final del libro
                      <span className="block text-[10px] font-normal opacity-75 mt-0.5">Después de la hoja {paginas.length}</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setModoInsercionJson('despues_actual')}
                      className={`p-3 rounded-xl border text-xs font-bold text-left transition-all cursor-pointer ${modoInsercionJson === 'despues_actual' ? 'border-blue-600 bg-blue-50 text-blue-900' : 'border-slate-200 text-slate-600 hover:bg-slate-50'}`}
                    >
                      📍 Después de hoja actual
                      <span className="block text-[10px] font-normal opacity-75 mt-0.5">Después de la hoja {indicePaginaActual + 1}</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setModoInsercionJson('reemplazar')}
                      className={`p-3 rounded-xl border text-xs font-bold text-left transition-all cursor-pointer ${modoInsercionJson === 'reemplazar' ? 'border-amber-600 bg-amber-50 text-amber-900' : 'border-slate-200 text-slate-600 hover:bg-slate-50'}`}
                    >
                      🔄 Reemplazar todas
                      <span className="block text-[10px] font-normal opacity-75 mt-0.5">Sobrescribe las actuales</span>
                    </button>
                  </div>
                </div>

                <div>
                  <label className="block text-xs font-bold text-slate-500 uppercase tracking-widest mb-2">Código JSON del Lote</label>
                  <textarea
                    rows={10}
                    value={jsonLoteInput}
                    onChange={(e) => setJsonLoteInput(e.target.value)}
                    placeholder={`[\n  {\n    "titulo": "Capítulo 1",\n    "subtitulo": "Estadísticas Nacionales",\n    "mostrarNumero": true,\n    "bloques": [\n      { "tipo": "texto", "contenido": "<p>Texto de la página...</p>" },\n      { "id": "tbl_1", "tipo": "tabla", "nombre": "Tabla 1", "data": [["Año", "Total"], ["2025", "100%"]] }\n    ]\n  }\n]`}
                    className="w-full p-4 bg-slate-900 text-emerald-300 font-mono text-xs rounded-2xl border border-slate-700 outline-none focus:border-emerald-500 custom-scrollbar"
                  />
                </div>
              </div>

              <div className="flex justify-between items-center mt-6 pt-4 border-t border-slate-100">
                <button
                  type="button"
                  onClick={() => setJsonLoteInput('')}
                  className="text-xs font-bold text-slate-400 hover:text-slate-600 cursor-pointer"
                >
                  Limpiar campo
                </button>
                <div className="flex gap-3">
                  <button
                    type="button"
                    onClick={() => setModalImportarJsonAbierto(false)}
                    className="px-5 py-2.5 rounded-xl font-bold text-slate-600 hover:bg-slate-100 text-sm transition-colors cursor-pointer"
                  >
                    Cancelar
                  </button>
                  <button
                    type="button"
                    onClick={procesarInyeccionLoteJson}
                    className="px-6 py-2.5 rounded-xl font-bold text-white bg-emerald-600 hover:bg-emerald-700 text-sm shadow-md transition-all flex items-center gap-2 cursor-pointer"
                  >
                    <CheckCircle2 size={16} /> Inyectar Páginas Ahora
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL PARA ELEGIR ENTRE CREAR TABLA EN BLANCO O CLONAR TABLA EXISTENTE */}
      {modalNuevaTablaAbierto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 no-print">
          <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm" onClick={() => setModalNuevaTablaAbierto(false)}></div>
          <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-xl relative z-10 animate-in zoom-in-95 duration-200 overflow-hidden">
            <div className="h-2 w-full bg-emerald-500"></div>
            <div className="p-8">
              <div className="flex justify-between items-start mb-4">
                <div>
                  <h3 className="text-xl font-black text-slate-900 tracking-tight flex items-center gap-2">
                    <TableIcon size={22} className="text-emerald-600" /> Insertar Bloque de Tabla
                  </h3>
                  <p className="text-xs text-slate-500 mt-1">Elige si deseas comenzar una tabla desde cero o clonar una tabla existente.</p>
                </div>
                <button onClick={() => setModalNuevaTablaAbierto(false)} className="text-slate-400 hover:text-slate-600 p-1 cursor-pointer"><X size={20} /></button>
              </div>

              <div className="space-y-6 mt-6">
                <button
                  type="button"
                  onClick={crearTablaEnBlanco}
                  className="w-full p-5 rounded-2xl border-2 border-emerald-500 bg-emerald-50/60 hover:bg-emerald-100/70 text-left transition-all flex items-center justify-between group cursor-pointer"
                >
                  <div>
                    <span className="block font-black text-emerald-950 text-base">Crear Tabla en Blanco desde Cero</span>
                    <span className="block text-xs text-emerald-700 mt-0.5">Inicia una tabla nueva estándar de 4 columnas lista para editar.</span>
                  </div>
                  <div className="p-3 bg-emerald-600 text-white rounded-xl group-hover:scale-105 transition-transform">
                    <Plus size={20} />
                  </div>
                </button>

                <div className="border-t border-slate-200 pt-5">
                  <h4 className="text-xs font-black uppercase tracking-widest text-slate-400 mb-3">
                    O Clonar / Duplicar una Tabla Existente ({todasLasTablasAdmin.length})
                  </h4>
                  {todasLasTablasAdmin.length === 0 ? (
                    <div className="p-6 text-center bg-slate-50 rounded-2xl border border-slate-200 text-xs text-slate-400">
                      Aún no hay otras tablas creadas en esta edición para clonar.
                    </div>
                  ) : (
                    <div className="max-h-64 overflow-y-auto space-y-2.5 pr-1">
                      {todasLasTablasAdmin.map((t) => (
                        <div
                          key={t.id}
                          className="p-3.5 rounded-xl border border-slate-200 hover:border-blue-400 bg-white flex items-center justify-between gap-4 transition-all"
                        >
                          <div className="min-w-0">
                            <p className="font-bold text-sm text-slate-800 truncate">{t.nombre || 'Tabla sin título'}</p>
                            <p className="text-[11px] text-slate-400">
                              Ubicación: {t.nombrePagina} • {t.data.length} filas × {t.data[0]?.length || 0} columnas
                            </p>
                          </div>
                          <button
                            type="button"
                            onClick={() => clonarTablaExistente(t)}
                            className="shrink-0 px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-lg flex items-center gap-1.5 shadow-xs transition-colors cursor-pointer"
                          >
                            <Copy size={13} /> Clonar esta
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              <div className="flex justify-end mt-6 pt-4 border-t border-slate-100">
                <button onClick={() => setModalNuevaTablaAbierto(false)} className="px-5 py-2.5 rounded-xl font-bold text-slate-600 hover:bg-slate-100 text-sm transition-colors cursor-pointer">
                  Cancelar
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {modalConfirmacion.abierto && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 no-print">
          <div className="absolute inset-0 bg-slate-900/40 backdrop-blur-sm" onClick={() => setModalConfirmacion(prev => ({ ...prev, abierto: false }))}></div>
          <div className="bg-white rounded-3xl shadow-2xl border border-slate-200 w-full max-w-md relative z-10 animate-in zoom-in-95 duration-200 overflow-hidden">
            <div className={`h-2 w-full ${modalConfirmacion.esPeligro ? 'bg-red-500' : 'bg-blue-500'}`}></div>
            <div className="p-8">
              <div className="flex justify-between items-start mb-4">
                <h3 className="text-xl font-black text-slate-900 tracking-tight">{modalConfirmacion.titulo}</h3>
                <button onClick={() => setModalConfirmacion(prev => ({ ...prev, abierto: false }))} className="text-slate-400 hover:text-slate-600 p-1 cursor-pointer"><X size={20} /></button>
              </div>
              <p className="text-slate-600 font-medium leading-relaxed mb-8">{modalConfirmacion.mensaje}</p>
              <div className="flex gap-3 justify-end">
                <button onClick={() => setModalConfirmacion(prev => ({ ...prev, abierto: false }))} className="px-5 py-3 rounded-xl font-bold text-slate-600 hover:bg-slate-100 transition-colors cursor-pointer">Cancelar</button>
                <button onClick={() => { if (modalConfirmacion.accionConfirmar) modalConfirmacion.accionConfirmar(); }} className={`px-5 py-3 rounded-xl font-bold text-white shadow-sm transition-all flex items-center gap-2 cursor-pointer ${modalConfirmacion.esPeligro ? 'bg-red-600 hover:bg-red-700' : 'bg-blue-600 hover:bg-blue-700'}`}>
                  {modalConfirmacion.esPeligro && <Trash2 size={16} />}
                  {modalConfirmacion.textoBotonConfirmar}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}