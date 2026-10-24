import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  Lock, Check, X, Download, ArrowRight, ArrowLeft, RefreshCw,
  Users, Trash2, Award, ClipboardCheck, ShieldAlert, RotateCcw,
  ChevronRight, Tag, Target, AlertCircle, Compass, Edit3, FileCheck
} from 'lucide-react';

const TEACHER_PASSWORD = 'doe2026';
const STORAGE_PREFIX = 'doe_taller_v6:';
const MIN_JUST = 20;

const C = {
  green: '#0F6E56', greenDeep: '#085041', greenBg: '#E1F5EE',
  blue: '#185FA5', blueBg: '#E6F1FB',
  amber: '#854F0B', amberBg: '#FAEEDA',
  coral: '#993C1D', coralBg: '#FAECE7',
  purple: '#534AB7', purpleBg: '#EEEDFE',
  gray: '#444441', grayBg: '#f5f5f3',
  paperBg: '#FAFAF6', cream: '#F5F8F4',
  ink: '#1a1a1a', muted: '#666', border: '#d0d0c8',
  disabled: '#bbb'
};

// =============== ARBOL DE DECISION ===============
const NODES = {
  start: {
    q: '¿Las variables independientes son proporciones que deben sumar a un total fijo?',
    note: 'La restricción Σxᵢ = 1 cambia la familia de diseños aplicable.',
    options: [
      { label: 'Factores que se ajustan de forma independiente', next: 'count_factors' },
      { label: 'Proporciones de una mezcla con suma constante', next: 'mixture_branch' }
    ]
  },
  mixture_branch: {
    q: '¿Hay variables de proceso (T, pH, tiempo) además de las proporciones?',
    note: 'Una mezcla con variables de proceso se estudia con un diseño cruzado.',
    options: [
      { label: 'Solamente las proporciones de los componentes', next: 'mixture_limits' },
      { label: 'Proporciones y una variable de proceso (T, pH)', leaf: 'mixture_process' }
    ]
  },
  mixture_limits: {
    q: '¿Cada componente puede tomar cualquier proporción entre 0 y 1?',
    note: 'Un límite inferior en los componentes reduce la región posible a un triángulo interior del simplex.',
    options: [
      { label: 'Cada componente puede ir de 0 a 1 sin límites', leaf: 'simplex' },
      { label: 'Cada componente tiene una proporción mínima', leaf: 'mixture_pseudo' }
    ]
  },
  count_factors: {
    q: '¿Cuántos factores controlables hay en el experimento?',
    note: 'El número de factores determina la familia de diseños aplicable.',
    options: [
      { label: 'Un solo factor', next: 'one_factor' },
      { label: '2 o 3 factores', next: 'few_factors' },
      { label: '4 o más factores', next: 'many_factors' }
    ]
  },
  one_factor: {
    q: '¿Existe una fuente de variabilidad sistemática (lotes, días, máquinas, jueces) que controlar?',
    note: 'Una fuente sistemática es candidata a factor de bloque para absorber variabilidad.',
    options: [
      { label: 'Condiciones homogéneas sin fuente sistemática', leaf: 'anova_one' },
      { label: 'Hay una fuente de variabilidad identificable', leaf: 'rcbd' }
    ]
  },
  few_factors: {
    q: '¿Cuál es el objetivo experimental con esos factores?',
    note: 'El objetivo determina el modelo, lineal con interacciones, cuadrático o con niveles fijos.',
    options: [
      { label: 'Estimar efectos principales e interacciones (lineal)', leaf: 'factorial_2k' },
      { label: 'Encontrar las condiciones que optimizan la respuesta', next: 'region_position' },
      { label: 'Comparar tres niveles fijos de interés por factor', leaf: 'factorial_3k' }
    ]
  },
  region_position: {
    q: '¿Qué indican los datos sobre la ubicación de la región experimental respecto al óptimo?',
    note: 'La prueba de curvatura con puntos centrales decide si basta un plano o se necesita un modelo de segundo orden.',
    options: [
      { label: 'Curvatura no significativa, la respuesta sigue un plano', leaf: 'steepest' },
      { label: 'Curvatura significativa, la región contiene el óptimo', next: 'second_order' }
    ]
  },
  second_order: {
    q: '¿Qué restricción impone la región de operación al diseño de segundo orden?',
    note: 'La forma de la región decide entre un diseño simétrico y uno óptimo.',
    options: [
      { label: 'Ninguna, los factores pueden tomar niveles en ±α', leaf: 'ccd_rot' },
      { label: 'Ningún factor puede salir del intervalo de −1 a +1', leaf: 'ccd_face' },
      { label: 'Los vértices del cubo son inseguros o muy costosos', leaf: 'bbd' },
      { label: 'Una restricción elimina una esquina de la región', leaf: 'd_optimal' }
    ]
  },
  many_factors: {
    q: 'Con muchos factores el factorial completo es costoso. ¿Qué prioriza el experimento ahora?',
    note: 'El protocolo estándar ejecuta primero el cribado y después la optimización.',
    options: [
      { label: 'Identificar los factores activos con pocas corridas', leaf: 'fractional' },
      { label: 'Ajustar un modelo cuadrático con todos los factores', leaf: 'rsm_warning' }
    ]
  }
};

const DESIGNS = {
  anova_one: { name: 'ANOVA de un factor', eq: 'yᵢⱼ = μ + τᵢ + εᵢⱼ',
    note: 'Comparación de medias de a tratamientos con n réplicas por tratamiento.' },
  rcbd: { name: 'Diseño de bloques completos al azar', eq: 'yᵢⱼ = μ + τᵢ + βⱼ + εᵢⱼ',
    note: 'El bloque βⱼ absorbe la variabilidad sistemática del lote, día, juez u operario.' },
  factorial_2k: { name: 'Factorial 2^k completo', eq: 'y = β₀ + Σ βᵢxᵢ + Σ βᵢⱼxᵢxⱼ + ε',
    note: 'Cada factor se estudia en los niveles −1 y +1. El modelo estima efectos principales e interacciones.' },
  factorial_3k: { name: 'Factorial 3^k', eq: 'y = β₀ + Σ βᵢxᵢ + Σ βᵢᵢxᵢ² + Σ βᵢⱼxᵢxⱼ + ε',
    note: 'Cada factor toma tres niveles fijos. El diseño estima la curvatura sin puntos axiales.' },
  steepest: { name: 'Factorial 2^k con puntos centrales y máxima pendiente',
    eq: 'ŷ = b₀ + Σ bᵢxᵢ',
    note: 'El plano ajustado en unidades codificadas define la ruta. Cada factor avanza con el signo de bᵢ y en proporción a su magnitud. La ruta se detiene tras dos descensos consecutivos y el mejor punto pasa a ser el nuevo centro.' },
  ccd_rot: { name: 'Diseño central compuesto rotable',
    eq: 'y = β₀ + Σ βᵢxᵢ + Σ βᵢᵢxᵢ² + Σ βᵢⱼxᵢxⱼ + ε',
    note: 'El diseño combina el factorial 2^k, 2k puntos axiales a una distancia α = (2^k)^¼ y C₀ puntos centrales. Cada factor toma cinco niveles.' },
  ccd_face: { name: 'Diseño central compuesto centrado en las caras',
    eq: 'y = β₀ + Σ βᵢxᵢ + Σ βᵢᵢxᵢ² + Σ βᵢⱼxᵢxⱼ + ε',
    note: 'El diseño usa α = 1. Los puntos axiales quedan en el centro de cada cara, todas las corridas están entre −1 y +1 y cada factor toma tres niveles.' },
  bbd: { name: 'Diseño de Box y Behnken',
    eq: 'y = β₀ + Σ βᵢxᵢ + Σ βᵢᵢxᵢ² + Σ βᵢⱼxᵢxⱼ + ε',
    note: 'Los puntos se ubican en el punto medio de las aristas del cubo y en el centro, sin vértices ni puntos axiales. Cada factor toma tres niveles y el diseño requiere k ≥ 3.' },
  d_optimal: { name: 'Diseño D óptimo',
    eq: 'max |XᵀX| sobre los puntos candidatos',
    note: 'Un algoritmo elige entre los puntos candidatos de la región restringida el subconjunto que maximiza el determinante de XᵀX para el modelo propuesto.' },
  rsm_warning: { name: 'Superficie de respuesta directa con muchos factores', eq: 'Sin modelo recomendado',
    note: 'Con cuatro o más factores, la optimización sin cribado previo gasta corridas en factores inactivos. El procedimiento no se recomienda.' },
  fractional: { name: 'Factorial fraccionado 2^(k−p)',
    eq: 'y = β₀ + Σ βᵢxᵢ + interacciones (con confusión)',
    note: 'El diseño reduce 2^k a 2^(k−p) corridas. La resolución III, IV o V depende de los generadores elegidos.' },
  simplex: { name: 'Diseño simplex con modelo de Scheffé',
    eq: 'y = Σ βᵢxᵢ + Σ βᵢⱼxᵢxⱼ',
    note: 'El modelo de Scheffé no tiene intercepto ni términos cuadráticos puros, porque Σxᵢ = 1 los hace colineales.' },
  mixture_pseudo: { name: 'Simplex en pseudocomponentes',
    eq: "xᵢ' = (xᵢ − Lᵢ) / (1 − ΣLⱼ)",
    note: 'Los límites inferiores Lᵢ reducen la región a un triángulo interior. Los pseudocomponentes de Kurotori convierten ese triángulo en un simplex completo donde se aplica el diseño de Scheffé.' },
  mixture_process: { name: 'Diseño cruzado de mezcla y proceso',
    eq: 'y = f(x) · g(z)',
    note: 'En cada punto del diseño de las variables de proceso z se corre el diseño de mezclas completo de las proporciones x.' }
};

// =============== DIAGRAMAS SVG ===============
function DesignDiagram({ designKey }) {
  const g = C.green, gBg = C.greenBg, mu = C.muted;
  switch (designKey) {
    case 'anova_one':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <line x1="20" y1="115" x2="220" y2="115" stroke={C.gray} strokeWidth="1.5"/>
          {[60, 90, 50, 100].map((h, i) => (
            <g key={i}>
              <rect x={30 + i*45} y={115-h} width="32" height={h} fill={g} opacity="0.85" rx="2"/>
              <text x={46 + i*45} y="132" fontSize="11" textAnchor="middle" fill={mu}>
                {['τ₁','τ₂','τ₃','τ₄'][i]}
              </text>
            </g>
          ))}
        </svg>
      );
    case 'rcbd':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          {[0, 1, 2].map(b => (
            <g key={b} transform={`translate(${10 + b*78}, 0)`}>
              <rect x="0" y="22" width="70" height="98" fill="none" stroke={g} strokeDasharray="4,2" opacity="0.6"/>
              <text x="35" y="16" fontSize="10" textAnchor="middle" fill={mu}>Bloque {b+1}</text>
              {[55, 70, 45].map((h, i) => (
                <rect key={i} x={6 + i*22} y={120-h} width="18" height={h} fill={g} opacity="0.85" rx="2"/>
              ))}
            </g>
          ))}
        </svg>
      );
    case 'factorial_2k':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(50, 12)">
            <path d="M 50 0 L 130 0 L 130 60 L 50 60 Z" fill="none" stroke={g} opacity="0.45"/>
            <path d="M 0 50 L 80 50 L 80 110 L 0 110 Z" fill={gBg} stroke={g} strokeWidth="2"/>
            <line x1="0" y1="50" x2="50" y2="0" stroke={g} opacity="0.7"/>
            <line x1="80" y1="50" x2="130" y2="0" stroke={g} opacity="0.7"/>
            <line x1="0" y1="110" x2="50" y2="60" stroke={g} opacity="0.7"/>
            <line x1="80" y1="110" x2="130" y2="60" stroke={g} opacity="0.7"/>
            {[[0,50],[80,50],[0,110],[80,110],[50,0],[130,0],[50,60],[130,60]].map(([x,y],i) => (
              <circle key={i} cx={x} cy={y} r="5" fill={g}/>
            ))}
          </g>
          <text x="120" y="135" fontSize="10" textAnchor="middle" fill={mu}>2³ = 8 corridas</text>
        </svg>
      );
    case 'factorial_3k':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(80, 15)">
            <line x1="0" y1="0" x2="0" y2="100" stroke={C.gray}/>
            <line x1="0" y1="100" x2="100" y2="100" stroke={C.gray}/>
            {[0,1,2].flatMap(i => [0,1,2].map(j => (
              <circle key={`${i}-${j}`} cx={i*50} cy={j*50} r="6" fill={g}/>
            )))}
            <text x="50" y="125" fontSize="10" textAnchor="middle" fill={mu}>X₁ (3 niveles)</text>
          </g>
          <text x="55" y="65" fontSize="10" textAnchor="middle" fill={mu} transform="rotate(-90 55 65)">X₂ (3 niveles)</text>
        </svg>
      );
    case 'ccd_rot':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(70, 15)">
            <rect x="20" y="20" width="80" height="80" fill={gBg} stroke={g} strokeWidth="2"/>
            <line x1="60" y1="60" x2="60" y2="-5" stroke={g} strokeDasharray="3,2" opacity="0.5"/>
            <line x1="60" y1="60" x2="60" y2="125" stroke={g} strokeDasharray="3,2" opacity="0.5"/>
            <line x1="60" y1="60" x2="-5" y2="60" stroke={g} strokeDasharray="3,2" opacity="0.5"/>
            <line x1="60" y1="60" x2="125" y2="60" stroke={g} strokeDasharray="3,2" opacity="0.5"/>
            {[[20,20],[100,20],[20,100],[100,100]].map(([x,y],i) => (
              <circle key={`c${i}`} cx={x} cy={y} r="5" fill={g}/>
            ))}
            {[[60,-5],[60,125],[-5,60],[125,60]].map(([x,y],i) => (
              <circle key={`a${i}`} cx={x} cy={y} r="5" fill="white" stroke={g} strokeWidth="2"/>
            ))}
            <circle cx="60" cy="60" r="6" fill={g}/>
          </g>
          <text x="120" y="135" fontSize="10" textAnchor="middle" fill={mu}>Factorial + axiales + central</text>
        </svg>
      );
    case 'fractional':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(50, 12)">
            <path d="M 50 0 L 130 0 L 130 60 L 50 60 Z" fill="none" stroke={g} opacity="0.25" strokeDasharray="2,2"/>
            <path d="M 0 50 L 80 50 L 80 110 L 0 110 Z" fill="none" stroke={g} opacity="0.4" strokeDasharray="2,2"/>
            <line x1="0" y1="50" x2="50" y2="0" stroke={g} opacity="0.25" strokeDasharray="2,2"/>
            <line x1="80" y1="50" x2="130" y2="0" stroke={g} opacity="0.25" strokeDasharray="2,2"/>
            <line x1="0" y1="110" x2="50" y2="60" stroke={g} opacity="0.25" strokeDasharray="2,2"/>
            <line x1="80" y1="110" x2="130" y2="60" stroke={g} opacity="0.25" strokeDasharray="2,2"/>
            {[[0,50],[80,110],[50,0],[130,60]].map(([x,y],i) => (
              <circle key={`f${i}`} cx={x} cy={y} r="5" fill={g}/>
            ))}
            {[[80,50],[0,110],[130,0],[50,60]].map(([x,y],i) => (
              <circle key={`u${i}`} cx={x} cy={y} r="5" fill="none" stroke={g} strokeDasharray="1.5,1.5" opacity="0.5"/>
            ))}
          </g>
          <text x="120" y="135" fontSize="10" textAnchor="middle" fill={mu}>2^(3−1): solo media fracción</text>
        </svg>
      );
    case 'simplex':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(45, 8)">
            <polygon points="80,10 10,115 150,115" fill={gBg} stroke={g} strokeWidth="2"/>
            <circle cx="80" cy="10" r="6" fill={g}/>
            <circle cx="10" cy="115" r="6" fill={g}/>
            <circle cx="150" cy="115" r="6" fill={g}/>
            <circle cx="45" cy="63" r="5" fill={g}/>
            <circle cx="115" cy="63" r="5" fill={g}/>
            <circle cx="80" cy="115" r="5" fill={g}/>
            <circle cx="80" cy="80" r="6" fill="white" stroke={g} strokeWidth="2"/>
            <text x="80" y="6" fontSize="9" textAnchor="middle" fill={mu}>x₁=1</text>
            <text x="0" y="130" fontSize="9" fill={mu}>x₂=1</text>
            <text x="160" y="130" fontSize="9" textAnchor="end" fill={mu}>x₃=1</text>
          </g>
        </svg>
      );
    case 'mixture_process':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(35, 8)">
            <polygon points="55,8 12,52 98,52" fill={gBg} stroke={g}/>
            {[[55,8],[12,52],[98,52]].map(([x,y],i) => <circle key={`t${i}`} cx={x} cy={y} r="3.5" fill={g}/>)}
            <polygon points="55,72 12,116 98,116" fill={gBg} stroke={g}/>
            {[[55,72],[12,116],[98,116]].map(([x,y],i) => <circle key={`b${i}`} cx={x} cy={y} r="3.5" fill={g}/>)}
            <line x1="55" y1="8" x2="55" y2="72" stroke={g} strokeDasharray="2,2" opacity="0.6"/>
            <line x1="12" y1="52" x2="12" y2="116" stroke={g} strokeDasharray="2,2" opacity="0.6"/>
            <line x1="98" y1="52" x2="98" y2="116" stroke={g} strokeDasharray="2,2" opacity="0.6"/>
            <line x1="135" y1="20" x2="135" y2="115" stroke={C.gray} strokeWidth="1.5"/>
            <polygon points="135,115 132,108 138,108" fill={C.gray}/>
            <text x="142" y="22" fontSize="10" fill={mu}>T alta</text>
            <text x="142" y="118" fontSize="10" fill={mu}>T baja</text>
          </g>
        </svg>
      );
    case 'steepest':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          {[0, 1, 2, 3].map(i => (
            <line key={i} x1={40 + i*40} y1="130" x2={110 + i*40} y2="10" stroke={g} opacity="0.25"/>
          ))}
          <rect x="40" y="80" width="40" height="40" fill={gBg} stroke={g} strokeWidth="1.5"/>
          {[[40,80],[80,80],[40,120],[80,120]].map(([x,y],i) => <circle key={i} cx={x} cy={y} r="4" fill={g}/>)}
          <circle cx="60" cy="100" r="4" fill="white" stroke={g} strokeWidth="2"/>
          <line x1="60" y1="100" x2="200" y2="28" stroke={C.coral} strokeWidth="2"/>
          <polygon points="200,28 190,30 195,38" fill={C.coral}/>
          {[[95,82],[130,64],[165,46]].map(([x,y],i) => <circle key={`r${i}`} cx={x} cy={y} r="3.5" fill={C.coral}/>)}
          <text x="150" y="128" fontSize="10" textAnchor="middle" fill={mu}>Ruta perpendicular a los contornos</text>
        </svg>
      );
    case 'ccd_face':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(70, 15)">
            <rect x="0" y="0" width="100" height="100" fill={gBg} stroke={g} strokeWidth="2"/>
            {[[0,0],[100,0],[0,100],[100,100]].map(([x,y],i) => <circle key={`c${i}`} cx={x} cy={y} r="5" fill={g}/>)}
            {[[50,0],[50,100],[0,50],[100,50]].map(([x,y],i) => <circle key={`a${i}`} cx={x} cy={y} r="5" fill="white" stroke={g} strokeWidth="2"/>)}
            <circle cx="50" cy="50" r="6" fill={g}/>
          </g>
          <text x="120" y="134" fontSize="10" textAnchor="middle" fill={mu}>α = 1, axiales en las caras</text>
        </svg>
      );
    case 'bbd':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(50, 12)">
            <path d="M 50 0 L 130 0 L 130 60 L 50 60 Z" fill="none" stroke={g} opacity="0.35"/>
            <path d="M 0 50 L 80 50 L 80 110 L 0 110 Z" fill="none" stroke={g} opacity="0.6"/>
            <line x1="0" y1="50" x2="50" y2="0" stroke={g} opacity="0.5"/>
            <line x1="80" y1="50" x2="130" y2="0" stroke={g} opacity="0.5"/>
            <line x1="0" y1="110" x2="50" y2="60" stroke={g} opacity="0.5"/>
            <line x1="80" y1="110" x2="130" y2="60" stroke={g} opacity="0.5"/>
            {[[90,0],[130,30],[90,60],[50,30],[40,50],[80,80],[40,110],[0,80],[25,25],[105,25],[25,85],[105,85]].map(([x,y],i) => (
              <circle key={i} cx={x} cy={y} r="4.5" fill={g}/>
            ))}
            <circle cx="65" cy="55" r="5" fill="white" stroke={g} strokeWidth="2"/>
          </g>
          <text x="120" y="136" fontSize="10" textAnchor="middle" fill={mu}>Puntos medios de las aristas, sin vértices</text>
        </svg>
      );
    case 'd_optimal':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(70, 15)">
            <polygon points="0,0 60,0 100,40 100,100 0,100" fill={gBg} stroke={g} strokeWidth="2"/>
            <polygon points="60,0 100,0 100,40" fill="none" stroke={C.coral} strokeDasharray="3,2"/>
            {[[0,0],[60,0],[100,40],[100,100],[0,100],[50,100],[0,50],[80,20],[50,50]].map(([x,y],i) => (
              <circle key={i} cx={x} cy={y} r="4.5" fill={g}/>
            ))}
          </g>
          <text x="120" y="134" fontSize="10" textAnchor="middle" fill={mu}>Región con una esquina prohibida</text>
        </svg>
      );
    case 'mixture_pseudo':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(45, 8)">
            <polygon points="80,10 10,115 150,115" fill="none" stroke={g} opacity="0.5"/>
            <polygon points="80,40 30,103 130,103" fill={gBg} stroke={g} strokeWidth="2"/>
            {[[80,40],[30,103],[130,103],[55,71],[105,71],[80,103],[80,82]].map(([x,y],i) => (
              <circle key={i} cx={x} cy={y} r="4.5" fill={g}/>
            ))}
            <text x="80" y="6" fontSize="9" textAnchor="middle" fill={mu}>x₁ = 1</text>
            <text x="80" y="132" fontSize="9" textAnchor="middle" fill={mu}>Región con límites inferiores</text>
          </g>
        </svg>
      );
    case 'rsm_warning':
      return (
        <svg viewBox="0 0 240 140" style={{ width: '100%', maxWidth: 280 }}>
          <g transform="translate(85, 30)">
            <circle cx="35" cy="40" r="32" fill={C.amberBg} stroke={C.amber} strokeWidth="2"/>
            <text x="35" y="55" fontSize="42" textAnchor="middle" fill={C.amber} fontWeight="700">!</text>
          </g>
          <text x="120" y="120" fontSize="10" textAnchor="middle" fill={C.amber}>Procedimiento no recomendado</text>
        </svg>
      );
    default:
      return null;
  }
}

// =============== CASOS DETECTIVE (8) ===============
const CASES = [
  {
    id: 'c1', area: 'Ingeniería ambiental',
    title: 'Caso 1 · Coagulantes en una planta de potabilización',
    scenario: 'Una planta de potabilización compara 4 coagulantes midiendo la turbidez residual del agua tratada en una prueba de jarras. El agua cruda cambia de un día a otro por las lluvias, y la turbidez inicial de cada día afecta de forma sistemática la eficiencia de todos los coagulantes. El laboratorio dispone de 5 días de ensayo y prueba los 4 coagulantes cada día, para un total de 20 corridas.',
    expected: 'rcbd', expectedPath: ['count_factors', 'one_factor']
  },
  {
    id: 'c2', area: 'Biotecnología',
    title: 'Caso 2 · Producción de un pigmento por levadura',
    scenario: 'Una planta produce un pigmento natural por fermentación con levadura a T = 30 °C y pH = 5.5. El equipo corrió un factorial 2² alrededor de esa condición con 4 puntos centrales. Los dos efectos principales resultaron significativos y la prueba de curvatura dio p = 0.47. El equipo quiere desplazar la operación hacia la región donde la concentración de pigmento sea mayor.',
    expected: 'steepest', expectedPath: ['count_factors', 'few_factors', 'region_position']
  },
  {
    id: 'c3', area: 'Ingeniería de alimentos',
    title: 'Caso 3 · Formulación de yogur firme',
    scenario: 'Un equipo de I+D formula el sistema estabilizante de un yogur con tres ingredientes, gelatina, almidón modificado y pectina, cuyas proporciones suman 1. Cada ingrediente puede ocupar cualquier fracción del sistema, desde 0 hasta la totalidad. El proceso de fermentación se mantiene constante y el objetivo es maximizar la firmeza del gel.',
    expected: 'simplex', expectedPath: ['mixture_branch', 'mixture_limits']
  },
  {
    id: 'c4', area: 'Ingeniería de materiales',
    title: 'Caso 4 · Inyección de piezas de polipropileno',
    scenario: 'Una empresa inyecta tapas de polipropileno y busca minimizar la contracción de la pieza con tres factores, temperatura de fundido, presión de sostenimiento y tiempo de enfriamiento. Un factorial con puntos centrales mostró curvatura significativa. La ficha técnica del material y los límites de la máquina impiden operar fuera de los niveles −1 y +1 del factorial, y todas las combinaciones dentro de ese cubo son operables.',
    expected: 'ccd_face', expectedPath: ['count_factors', 'few_factors', 'region_position', 'second_order']
  },
  {
    id: 'c5', area: 'Ingeniería química',
    title: 'Caso 5 · Extracción de aceite esencial',
    scenario: 'Un ingeniero estudia la extracción de aceite esencial de cáscara de naranja con 6 factores controlables, temperatura, tiempo, relación de solvente, agitación, tamaño de partícula y humedad inicial. Sospecha que solo 2 o 3 factores son activos y cuenta con un presupuesto máximo de 16 corridas para identificarlos.',
    expected: 'fractional', expectedPath: ['count_factors', 'many_factors']
  },
  {
    id: 'c6', area: 'Industria farmacéutica',
    title: 'Caso 6 · Síntesis de un principio activo',
    scenario: 'Un laboratorio farmacéutico optimiza el rendimiento de una síntesis con tres factores, temperatura, concentración del catalizador y tiempo de reacción. Un factorial previo con puntos centrales mostró curvatura significativa. La combinación simultánea de los tres niveles altos produce una reacción exotérmica descontrolada, y la de los tres niveles bajos no genera conversión medible.',
    expected: 'bbd', expectedPath: ['count_factors', 'few_factors', 'region_position', 'second_order']
  },
  {
    id: 'c7', area: 'Ingeniería de alimentos',
    title: 'Caso 7 · Secado por aspersión de extracto de café',
    scenario: 'Una planta de café soluble ajusta un modelo cuadrático de la humedad del polvo con dos factores, temperatura del aire de entrada (160 a 220 °C) y caudal de alimentación (10 a 30 mL/min). Un ensayo previo con puntos centrales mostró curvatura significativa. Con temperatura mayor de 200 °C y caudal menor de 15 mL/min el polvo se quema, así que esa esquina de la región queda prohibida.',
    expected: 'd_optimal', expectedPath: ['count_factors', 'few_factors', 'region_position', 'second_order']
  },
  {
    id: 'c8', area: 'Ingeniería civil',
    title: 'Caso 8 · Aglomerante con cenizas y escoria',
    scenario: 'Una cementera formula un aglomerante con cemento x₁, ceniza volante x₂ y escoria x₃, con x₁ + x₂ + x₃ = 1. La norma exige al menos 0.5 de cemento, y la planta exige al menos 0.1 de ceniza y 0.1 de escoria. El curado se hace en condiciones normalizadas fijas y el objetivo es maximizar la resistencia a compresión a 28 días.',
    expected: 'mixture_pseudo', expectedPath: ['mixture_branch', 'mixture_limits']
  }
];

// =============== AUDITORIAS (5) ===============
const AUDIT_TAGS = {
  replicacion: { label: 'Falta replicación', color: C.coral },
  aleatorizacion: { label: 'Falta aleatorización', color: C.amber },
  modelo: { label: 'Modelo incorrecto', color: C.purple },
  prueba: { label: 'Prueba estadística incorrecta', color: C.blue },
  curvatura: { label: 'Curvatura ignorada', color: C.green },
  escala: { label: 'Uso incorrecto de unidades codificadas', color: C.blue },
  falta_ajuste: { label: 'Falta de ajuste ignorada', color: C.purple },
  extrapolacion: { label: 'Extrapolación injustificada', color: C.coral },
  restriccion: { label: 'Restricción de mezcla ignorada', color: C.green },
  conclusion: { label: 'Conclusión injustificada por los datos', color: C.amber }
};

const AUDIT_CASES = [
  {
    id: 'a1', area: 'Ingeniería de alimentos',
    title: 'Auditoría A · Extracción de pectina',
    context: {
      objetivo: 'Maximizar el rendimiento de extracción de pectina a partir de cáscara de limón.',
      factores: 'A, temperatura (70 °C, 90 °C). B, pH (1.5, 2.5). C, tiempo (60 min, 120 min).',
      respuesta: 'Rendimiento gravimétrico de pectina seca (g de pectina por 100 g de cáscara seca).',
      diseno: 'Factorial 2³ planeado, 8 corridas únicas.'
    },
    intro: 'Lee el extracto del informe del equipo. Haz clic en cada frase con error metodológico y asígnale la categoría desde el panel lateral.',
    spans: [
      { id: 's1', text: 'El protocolo fijó los niveles de A, B y C a partir de corridas preliminares, y el equipo ejecutó las 8 combinaciones del diseño 2³' },
      { id: 's2', text: '. ' },
      { id: 's3', text: 'Las 8 corridas se hicieron sin réplicas y sin usar las interacciones de alto orden como estimación del error experimental', error: 'replicacion' },
      { id: 's4', text: '. ' },
      { id: 's5', text: 'Para la significancia aplicó una prueba t entre las 4 corridas con A alto y las 4 con A bajo, y reportó p = 0.03 para A', error: 'prueba' },
      { id: 's6', text: '. ' },
      { id: 's7', text: 'Las corridas se ejecutaron en el orden estándar de Yates (− − −, − − +, − + −, ...) en que aparecen en la matriz de diseño', error: 'aleatorizacion' },
      { id: 's8', text: '. ' },
      { id: 's9', text: 'La discusión concluye que solo A es activo, sin examinar AB, AC, BC ni ABC en una gráfica de probabilidad normal de efectos', error: 'modelo' },
      { id: 's10', text: '.' }
    ]
  },
  {
    id: 'a2', area: 'Biotecnología',
    title: 'Auditoría B · OVAT y máxima pendiente en un biosurfactante',
    context: {
      objetivo: 'Maximizar la concentración de un biosurfactante producido por fermentación.',
      factores: 'T, temperatura (40 °C a 48 °C). t, tiempo de incubación (18 h a 30 h). La planta opera a 44 °C y 24 h.',
      respuesta: 'Concentración de biosurfactante al final del lote (g/L).',
      diseno: 'Estrategia OVAT inicial, factorial 2² con 3 puntos centrales y ruta de máxima pendiente.'
    },
    intro: 'Lee el extracto del informe. Identifica los errores de la estrategia OVAT y de la aproximación al óptimo.',
    spans: [
      { id: 's1', text: 'El equipo varió primero la temperatura con el tiempo fijo en 24 h, y después el tiempo con la temperatura fija en el mejor valor encontrado en el barrido' },
      { id: 's2', text: '. ' },
      { id: 's3', text: 'El informe declara esa combinación como el óptimo del proceso, aunque el factorial posterior mostró una interacción T × t significativa', error: 'conclusion' },
      { id: 's4', text: '. ' },
      { id: 's5', text: 'Los puntos centrales mostraron curvatura significativa (p = 0.01), y el equipo trazó de todos modos la ruta de máxima pendiente con el plano ajustado', error: 'curvatura' },
      { id: 's6', text: '. ' },
      { id: 's7', text: 'La dirección de la ruta salió de los coeficientes del modelo en unidades reales, °C y h, y cada paso movió los factores en esa proporción', error: 'escala' },
      { id: 's8', text: '. ' },
      { id: 's9', text: 'La ruta avanzó con un experimento en cada paso, y el recorrido se detuvo después de dos descensos consecutivos de la concentración' },
      { id: 's10', text: '.' }
    ]
  },
  {
    id: 'a3', area: 'Ingeniería de alimentos',
    title: 'Auditoría C · Mezcla de grasas en margarina',
    context: {
      objetivo: 'Maximizar el punto de fusión de una margarina manteniendo su perfil sensorial.',
      factores: 'Tres grasas vegetales x₁ (palma), x₂ (soya) y x₃ (algodón), con x₁ + x₂ + x₃ = 1.',
      respuesta: 'Punto de fusión slip (°C).',
      diseno: 'Simplex centroide aumentado con 10 puntos.'
    },
    intro: 'Lee el extracto. Identifica errores de modelo, restricción y conclusión.',
    spans: [
      { id: 's0', text: 'El diseño incluyó los tres componentes puros, las tres mezclas binarias 1/2, el centroide y tres mezclas axiales en el interior del triángulo' },
      { id: 's0b', text: '. ' },
      { id: 's1', text: 'El investigador planteó como modelo y = β₀ + β₁x₁ + β₂x₂ + β₃x₃ + β₁₂x₁x₂ + β₁₃x₁x₃ + β₂₃x₂x₃ + β₁₁x₁² + β₂₂x₂² + β₃₃x₃²', error: 'modelo' },
      { id: 's2', text: '. ' },
      { id: 's3', text: 'Reportó intervalos de confianza individuales para β₀, β₁, β₂ y β₃, suponiendo coeficientes independientes y una matriz de diseño no singular', error: 'restriccion' },
      { id: 's4', text: '. ' },
      { id: 's5', text: 'Con un R² ajustado mayor de 0.99, concluyó que la palma es el componente más relevante porque β₁ fue el mayor término lineal', error: 'conclusion' },
      { id: 's6', text: '.' }
    ]
  },
  {
    id: 'a4', area: 'Ingeniería química',
    title: 'Auditoría D · Reactor catalítico con DCC',
    context: {
      objetivo: 'Maximizar el rendimiento de un reactor catalítico de hidrogenación selectiva.',
      factores: 'T, temperatura (60, 70, 80 °C). t, tiempo de residencia (20, 30, 40 min).',
      respuesta: 'Conversión a producto deseado (%).',
      diseno: 'DCC con 4 puntos factoriales, 4 axiales (α = 1.414) y 5 centrales, para un total de 13 corridas.'
    },
    intro: 'Lee el extracto del informe final. Identifica los errores de interpretación del modelo de segundo orden.',
    spans: [
      { id: 's1', text: 'El modelo cuadrático completo y el análisis canónico dieron un punto estacionario en T = 95 °C y t = 55 min, que los valores propios negativos clasifican como máximo' },
      { id: 's2', text: '. ' },
      { id: 's3', text: 'La prueba de falta de ajuste con los 5 puntos centrales dio p = 0.004, y el equipo usó el modelo cuadrático sin cambios para localizar el óptimo', error: 'falta_ajuste' },
      { id: 's4', text: '. ' },
      { id: 's5', text: 'El equipo reportó T = 95 °C y t = 55 min directamente como condición óptima recomendada para implementación industrial', error: 'extrapolacion' },
      { id: 's6', text: '. ' },
      { id: 's7', text: 'No se ejecutaron corridas de verificación experimental ni se desplazó el centro del diseño hacia la región del óptimo predicho antes de reportar la recomendación', error: 'conclusion' },
      { id: 's8', text: '.' }
    ]
  },
  {
    id: 'a5', area: 'Ingeniería civil',
    title: 'Auditoría E · Aglomerante con límites en los componentes',
    context: {
      objetivo: 'Maximizar la resistencia a compresión de un aglomerante de cemento, ceniza volante y escoria.',
      factores: 'x₁ (cemento) ≥ 0.5, x₂ (ceniza volante) ≥ 0.1 y x₃ (escoria) ≥ 0.1, con x₁ + x₂ + x₃ = 1.',
      respuesta: 'Resistencia a compresión a 28 días (MPa).',
      diseno: 'Simplex centroide de 7 puntos con 3 réplicas del centroide, para un total de 10 corridas.'
    },
    intro: 'Lee el extracto. Identifica los errores en el tratamiento de las restricciones y en la lectura del modelo.',
    spans: [
      { id: 's1', text: "Las proporciones se transformaron a pseudocomponentes con xᵢ' = (xᵢ − Lᵢ) / 0.3, y el simplex centroide se construyó sobre esas nuevas variables" },
      { id: 's2', text: '. ' },
      { id: 's3', text: 'El modelo de Scheffé se ajustó con las proporciones reales xᵢ y cada coeficiente lineal se reportó como la resistencia del componente puro', error: 'restriccion' },
      { id: 's4', text: '. ' },
      { id: 's5', text: 'Las 3 réplicas del centroide dieron una prueba de falta de ajuste significativa (p = 0.02), y el equipo mantuvo el modelo cuadrático para predecir', error: 'falta_ajuste' },
      { id: 's6', text: '. ' },
      { id: 's7', text: 'Como b₁ fue el mayor coeficiente lineal, el informe concluyó que más cemento sube la resistencia en toda la región, sin la gráfica de trazas', error: 'conclusion' },
      { id: 's8', text: '.' }
    ]
  }
];

// =============== COMPONENTE PRINCIPAL ===============
export default function TallerDOE() {
  const [mode, setMode] = useState('home');
  return (
    <div style={{
      minHeight: '100vh',
      background: `linear-gradient(135deg, ${C.paperBg} 0%, ${C.cream} 100%)`,
      fontFamily: '"Crimson Pro", Georgia, serif', color: C.ink
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Crimson+Pro:wght@400;500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');
        .font-serif { font-family: "Crimson Pro", Georgia, serif; }
        .font-sans { font-family: "Inter", system-ui, sans-serif; }
        .font-mono { font-family: "JetBrains Mono", monospace; }
        @keyframes fadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
        .fade-in { animation: fadeIn 0.3s ease-out; }
      `}</style>
      <div className="max-w-5xl mx-auto px-6 py-8">
        {mode === 'home' && <Home onSelect={setMode} />}
        {mode === 'student' && <StudentFlow onExit={() => setMode('home')} />}
        {mode === 'teacher' && <TeacherFlow onExit={() => setMode('home')} />}
      </div>
    </div>
  );
}

// =============== HOME ===============
function Home({ onSelect }) {
  return (
    <div className="min-h-screen flex flex-col">
      <div className="flex-1 flex flex-col items-center justify-center text-center py-16">
        <div className="inline-block px-4 py-1 mb-6 rounded-full font-sans text-xs uppercase tracking-widest"
          style={{ background: C.green, color: '#fff' }}>
          Diseño de Experimentos · Tercer corte
        </div>
        <h1 className="font-serif font-bold mb-6" style={{
          color: C.ink, fontSize: 'clamp(2.5rem, 6vw, 4.5rem)', lineHeight: 1.05
        }}>
          Taller integrador
        </h1>
        <p className="text-base font-serif max-w-2xl mb-10" style={{ color: C.ink, lineHeight: 1.7, fontSize: 17 }}>
          Este taller integra los temas del semestre, desde la comparación de tratamientos
          hasta la superficie de respuesta y los diseños de mezclas. Cada caso plantea una
          decisión real de la industria o de la academia, en la elección del diseño o en la
          lectura crítica de un informe experimental.
        </p>
        <button onClick={() => onSelect('student')}
          className="group px-10 py-5 rounded-full font-sans font-semibold text-lg flex items-center gap-3"
          style={{
            background: C.green, color: '#fff', cursor: 'pointer', border: 'none',
            transition: 'transform .2s, box-shadow .2s'
          }}
          onMouseEnter={e => { e.currentTarget.style.transform = 'scale(1.05)'; e.currentTarget.style.boxShadow = '0 10px 30px rgba(15,110,86,0.3)'; }}
          onMouseLeave={e => { e.currentTarget.style.transform = 'scale(1)'; e.currentTarget.style.boxShadow = 'none'; }}>
          Comenzar taller
          <ArrowRight size={22} />
        </button>
        <div className="mt-14 grid grid-cols-2 gap-6 max-w-2xl">
          <FeatureCard icon={Compass} color={C.green}
            title="Parte 1 · DOE Detective"
            text={`${CASES.length} escenarios. Recorres un árbol de preguntas diagnósticas y el sistema muestra el diseño que tus respuestas implican. Cada caso exige una justificación técnica.`} />
          <FeatureCard icon={Target} color={C.coral}
            title="Parte 2 · Auditoría"
            text={`${AUDIT_CASES.length} informes con errores metodológicos. Marcas cada frase con error y le asignas su categoría. Cada informe incluye su contexto experimental.`} />
        </div>
      </div>
      <div className="text-center pb-4">
        <button onClick={() => onSelect('teacher')}
          style={{
            background: 'transparent', border: 'none', cursor: 'pointer',
            color: C.muted, fontSize: 11, opacity: 0.5,
            display: 'inline-flex', alignItems: 'center', gap: 4
          }}
          onMouseEnter={e => e.currentTarget.style.opacity = 1}
          onMouseLeave={e => e.currentTarget.style.opacity = 0.5}>
          <Lock size={11} /> Acceso docente
        </button>
      </div>
    </div>
  );
}

function FeatureCard({ icon: Icon, color, title, text }) {
  return (
    <div className="p-5 rounded-xl text-left" style={{
      background: 'rgba(255,255,255,0.7)', border: `1px solid ${color}`, borderLeft: `4px solid ${color}`
    }}>
      <Icon size={22} style={{ color }} className="mb-2" />
      <div className="font-serif font-bold text-lg mb-1" style={{ color: C.ink }}>{title}</div>
      <div className="font-sans text-sm leading-relaxed" style={{ color: C.muted }}>{text}</div>
    </div>
  );
}

// =============== FLUJO ESTUDIANTE ===============
function StudentFlow({ onExit }) {
  const [step, setStep] = useState('auth');
  const [student, setStudent] = useState({ name: '', cedula: '', program: '' });
  const [detectiveAnswers, setDetectiveAnswers] = useState({});
  const [auditAnswers, setAuditAnswers] = useState({});
  const [section, setSection] = useState('detective');
  const [idx, setIdx] = useState(0);
  const [submitting, setSubmitting] = useState(false);
  const [finalScore, setFinalScore] = useState(null);

  const total = CASES.length + AUDIT_CASES.length;
  const globalIdx = section === 'detective' ? idx : CASES.length + idx;
  const isLast = section === 'audit' && idx === AUDIT_CASES.length - 1;
  const currentCase = section === 'detective' ? CASES[idx] : AUDIT_CASES[idx];
  const currentAns = section === 'detective' ? detectiveAnswers[currentCase.id] : auditAnswers[currentCase.id];

  const goNext = () => {
    if (section === 'detective') {
      if (idx < CASES.length - 1) setIdx(idx + 1);
      else { setSection('audit'); setIdx(0); }
    } else if (idx < AUDIT_CASES.length - 1) setIdx(idx + 1);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const goPrev = () => {
    if (section === 'audit' && idx === 0) {
      setSection('detective'); setIdx(CASES.length - 1);
    } else if (idx > 0) setIdx(idx - 1);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const goToCase = (sectionTarget, idxTarget) => {
    setSection(sectionTarget); setIdx(idxTarget);
    setStep('workshop');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const computeScore = () => {
    let detPoints = 0;
    const detDetail = {};
    CASES.forEach(c => {
      const ans = detectiveAnswers[c.id];
      if (!ans) {
        detDetail[c.id] = { points: 0, max: 1, leaf: null, expected: c.expected, path: [], justification: '' };
        return;
      }
      const reachedCorrect = ans.leaf === c.expected;
      const fullPath = [...c.expectedPath, c.expected];
      const userPath = ans.path.map(p => p.next || p.leaf);
      let correctSteps = 0;
      for (let i = 0; i < Math.min(userPath.length, fullPath.length); i++) {
        if (userPath[i] === fullPath[i]) correctSteps++;
        else break;
      }
      const points = reachedCorrect ? 1 : (correctSteps / fullPath.length);
      detPoints += points;
      detDetail[c.id] = { points, max: 1, leaf: ans.leaf, expected: c.expected, path: ans.path, justification: ans.justification || '' };
    });

    let audPoints = 0;
    const audDetail = {};
    AUDIT_CASES.forEach(c => {
      const ans = auditAnswers[c.id] || { tags: {} };
      const errorSpans = c.spans.filter(s => s.error);
      let correctTags = 0, wrongTags = 0;
      Object.entries(ans.tags).forEach(([spanId, tag]) => {
        const span = c.spans.find(s => s.id === spanId);
        if (!span) return;
        if (span.error === tag) correctTags++;
        else wrongTags++;
      });
      const E = errorSpans.length;
      const raw = (correctTags / E) - (0.5 * wrongTags / E);
      const pts = Math.max(0, Math.min(1, raw));
      audPoints += pts;
      audDetail[c.id] = { points: pts, max: 1, correctTags, wrongTags, totalErrors: E, userTags: ans.tags };
    });

    const auto = detPoints + audPoints;
    const totalMax = CASES.length + AUDIT_CASES.length;
    return { auto, totalMax, detPoints, audPoints, detDetail, audDetail, scaled: (auto / totalMax) * 5 };
  };

  const handleFinalSubmit = async () => {
    setSubmitting(true);
    const scoreData = computeScore();
    const submission = {
      ...student, detectiveAnswers, auditAnswers,
      score: scoreData, submittedAt: new Date().toISOString()
    };
    try {
      await window.storage.set(`${STORAGE_PREFIX}${student.cedula.trim()}`,
        JSON.stringify(submission), true);
      setFinalScore(scoreData);
      setStep('done');
    } catch (err) {
      alert('Error al guardar: ' + err.message);
    }
    setSubmitting(false);
  };

  const allDone = CASES.every(c => detectiveAnswers[c.id]?.confirmed) &&
    AUDIT_CASES.every(c => auditAnswers[c.id]?.confirmed);

  if (step === 'auth') return <AuthScreen student={student} setStudent={setStudent}
    onContinue={() => setStep('workshop')} onExit={onExit} />;

  if (step === 'review') return <ReviewScreen
    student={student} detectiveAnswers={detectiveAnswers} auditAnswers={auditAnswers}
    onEdit={goToCase} onSubmit={handleFinalSubmit} onBackToWorkshop={() => setStep('workshop')}
    submitting={submitting} />;

  if (step === 'done') return <DoneScreen student={student} score={finalScore} onExit={onExit} />;

  return (
    <div className="fade-in">
      <Progress current={globalIdx} total={total}
        detectiveAnswers={detectiveAnswers} auditAnswers={auditAnswers} />
      {section === 'detective' ? (
        <DetectiveCase key={currentCase.id} caseData={currentCase} answer={currentAns}
          onAnswer={a => setDetectiveAnswers(prev => ({ ...prev, [currentCase.id]: a }))} />
      ) : (
        <AuditCase key={currentCase.id} caseData={currentCase} answer={currentAns}
          onAnswer={a => setAuditAnswers(prev => ({ ...prev, [currentCase.id]: a }))} />
      )}
      <div className="flex justify-between mt-6 gap-3">
        <button onClick={() => { if (!(section === 'detective' && idx === 0)) goPrev(); }}
          style={{
            padding: '12px 24px', borderRadius: 8, fontWeight: 600, fontSize: 14,
            background: '#fff', color: C.ink, border: `2px solid ${C.ink}`,
            cursor: section === 'detective' && idx === 0 ? 'not-allowed' : 'pointer',
            opacity: section === 'detective' && idx === 0 ? 0.3 : 1,
            display: 'flex', alignItems: 'center', gap: 8, fontFamily: 'Inter, sans-serif'
          }}>
          <ArrowLeft size={18} /> Anterior
        </button>
        {!isLast ? (
          <button onClick={goNext}
            style={{
              padding: '12px 24px', borderRadius: 8, fontWeight: 600, fontSize: 14,
              background: C.green, color: '#fff', border: 'none', cursor: 'pointer',
              display: 'flex', alignItems: 'center', gap: 8, fontFamily: 'Inter, sans-serif'
            }}>
            Siguiente caso <ArrowRight size={18} />
          </button>
        ) : (
          <button onClick={() => { if (allDone) setStep('review'); }}
            style={{
              padding: '12px 24px', borderRadius: 8, fontWeight: 600, fontSize: 14,
              background: allDone ? C.coral : C.disabled, color: '#fff', border: 'none',
              cursor: allDone ? 'pointer' : 'not-allowed', opacity: allDone ? 1 : 0.5,
              display: 'flex', alignItems: 'center', gap: 8, fontFamily: 'Inter, sans-serif'
            }}>
            Revisar antes de entregar <FileCheck size={18} />
          </button>
        )}
      </div>
      {!allDone && isLast && (
        <div className="mt-4 p-4 rounded-lg flex items-center gap-2"
          style={{
            background: C.amberBg, color: C.amber, border: `1px solid ${C.amber}`,
            fontFamily: 'Inter, sans-serif', fontSize: 13
          }}>
          <AlertCircle size={18} />
          Faltan casos por confirmar. Cada caso debe terminar con su botón de confirmar antes de poder revisar.
        </div>
      )}
    </div>
  );
}

// =============== AUTENTICACION ===============
function AuthScreen({ student, setStudent, onContinue, onExit }) {
  const valid = student.name.trim() && student.cedula.trim();
  return (
    <div className="max-w-md mx-auto py-12 fade-in">
      <button onClick={onExit} style={{
        background: 'none', border: 'none', cursor: 'pointer',
        color: C.muted, fontSize: 14, marginBottom: 24,
        display: 'flex', alignItems: 'center', gap: 4, fontFamily: 'Inter, sans-serif'
      }}>
        <ArrowLeft size={14} /> Volver
      </button>
      <div className="bg-white rounded-2xl p-8 shadow-lg" style={{ border: `1px solid ${C.green}` }}>
        <h2 className="text-3xl font-serif font-bold mb-2" style={{ color: C.ink }}>Identificación</h2>
        <p className="font-sans text-sm mb-6" style={{ color: C.muted }}>
          Tu cédula sirve como identificador único. Una nueva entrega con la misma cédula reemplaza la anterior.
        </p>
        <div className="space-y-4">
          <Field label="Nombre completo" value={student.name}
            onChange={v => setStudent({ ...student, name: v })}
            placeholder="Ej. María Fernanda García" />
          <Field label="Cédula" value={student.cedula}
            onChange={v => setStudent({ ...student, cedula: v.replace(/\s/g, '') })}
            placeholder="Ej. 1014123456" />
          <Field label="Programa (opcional)" value={student.program}
            onChange={v => setStudent({ ...student, program: v })}
            placeholder="Ej. Ingeniería de Alimentos" />
        </div>
        <button onClick={() => { if (valid) onContinue(); }}
          style={{
            width: '100%', marginTop: 24, padding: '12px 24px', borderRadius: 8,
            fontWeight: 600, fontSize: 15, fontFamily: 'Inter, sans-serif',
            background: valid ? C.green : C.disabled, color: '#fff', border: 'none',
            cursor: valid ? 'pointer' : 'not-allowed', opacity: valid ? 1 : 0.5,
            display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 8
          }}>
          Comenzar <ArrowRight size={18} />
        </button>
      </div>
    </div>
  );
}

function Field({ label, value, onChange, placeholder, type = 'text' }) {
  return (
    <div>
      <label className="block font-sans text-xs font-semibold uppercase tracking-wider mb-1"
        style={{ color: C.muted }}>{label}</label>
      <input type={type} value={value} onChange={e => onChange(e.target.value)}
        placeholder={placeholder}
        className="w-full px-4 py-3 rounded-lg font-sans focus:outline-none"
        style={{ border: `1px solid ${C.border}`, background: C.cream }} />
    </div>
  );
}

// =============== PROGRESS ===============
function Progress({ current, total, detectiveAnswers, auditAnswers }) {
  return (
    <div className="mb-6">
      <div className="flex justify-between font-sans text-xs uppercase tracking-widest mb-2"
        style={{ color: C.muted }}>
        <span>Caso {current + 1} de {total}</span>
        <span>
          <span style={{ color: C.green }}>{CASES.filter(c => detectiveAnswers[c.id]?.confirmed).length}/{CASES.length} detective</span>
          {' · '}
          <span style={{ color: C.coral }}>{AUDIT_CASES.filter(c => auditAnswers[c.id]?.confirmed).length}/{AUDIT_CASES.length} auditoría</span>
        </span>
      </div>
      <div className="flex gap-1">
        {CASES.map((c, i) => {
          const done = detectiveAnswers[c.id]?.confirmed;
          const isCurrent = i === current;
          return (
            <div key={c.id} className="flex-1 h-2 rounded-full transition-all"
              style={{
                background: done ? C.green : C.border,
                opacity: isCurrent ? 1 : done ? 0.85 : 0.5,
                outline: isCurrent ? `2px solid ${C.green}` : 'none', outlineOffset: '2px'
              }} />
          );
        })}
        {AUDIT_CASES.map((c, i) => {
          const done = auditAnswers[c.id]?.confirmed;
          const isCurrent = (CASES.length + i) === current;
          return (
            <div key={c.id} className="flex-1 h-2 rounded-full transition-all"
              style={{
                background: done ? C.coral : C.border,
                opacity: isCurrent ? 1 : done ? 0.85 : 0.5,
                outline: isCurrent ? `2px solid ${C.coral}` : 'none', outlineOffset: '2px'
              }} />
          );
        })}
      </div>
    </div>
  );
}

// =============== CASO DETECTIVE (BOTON DE CONFIRMAR ARREGLADO) ===============
function DetectiveCase({ caseData, answer, onAnswer }) {
  const [path, setPath] = useState(answer?.path || []);
  const [nodeId, setNodeId] = useState(answer?.nodeId || 'start');
  const [leaf, setLeaf] = useState(answer?.leaf || null);
  const [justification, setJustification] = useState(answer?.justification || '');
  const [confirmed, setConfirmed] = useState(answer?.confirmed || false);

  const node = nodeId ? NODES[nodeId] : null;
  const justLength = justification.trim().length;
  const canConfirm = justLength >= MIN_JUST;

  const handleChoice = (optIdx) => {
    if (confirmed) return;
    const opt = node.options[optIdx];
    const step = { nodeId, label: opt.label, next: opt.next, leaf: opt.leaf };
    const newPath = [...path, step];
    setPath(newPath);
    if (opt.leaf) { setLeaf(opt.leaf); setNodeId(null); }
    else setNodeId(opt.next);
  };

  const handleReset = () => {
    setPath([]); setNodeId('start'); setLeaf(null);
    setJustification(''); setConfirmed(false);
  };

  // Sin atributo `disabled` para evitar quirks. Validación dentro del handler.
  const handleConfirmClick = () => {
    if (!canConfirm) return;
    setConfirmed(true);
    onAnswer({ path, leaf, nodeId, justification, confirmed: true });
  };

  return (
    <div className="bg-white rounded-2xl p-7 shadow-md fade-in" style={{ border: `1px solid ${C.green}` }}>
      <div className="flex items-center gap-2 mb-3 flex-wrap">
        <span className="font-sans text-xs uppercase tracking-widest px-3 py-1 rounded-full"
          style={{ background: C.green, color: '#fff' }}>
          <Compass size={11} style={{ display: 'inline', marginRight: 4 }} /> Parte 1 · Detective
        </span>
        <span className="font-sans text-xs italic" style={{ color: C.muted }}>{caseData.area}</span>
      </div>
      <h3 className="text-2xl font-serif font-bold mb-3" style={{ color: C.ink }}>{caseData.title}</h3>
      <div className="p-4 mb-5 rounded-lg font-serif text-base leading-relaxed"
        style={{ background: C.greenBg, borderLeft: `3px solid ${C.green}`, color: C.ink }}>
        {caseData.scenario}
      </div>

      {path.length > 0 && (
        <div className="mb-5">
          <div className="font-sans text-xs uppercase tracking-widest mb-2" style={{ color: C.muted }}>
            Tu ruta diagnóstica
          </div>
          <div className="flex flex-wrap items-center gap-2">
            {path.map((s, i) => (
              <React.Fragment key={i}>
                <div className="px-3 py-1 rounded-full font-sans text-xs font-semibold"
                  style={{ background: C.greenBg, color: C.green, border: `1px solid ${C.green}` }}>
                  {s.label}
                </div>
                {i < path.length - 1 && <ChevronRight size={12} style={{ color: C.muted }} />}
              </React.Fragment>
            ))}
          </div>
        </div>
      )}

      {node && !confirmed && (
        <div className="fade-in">
          <div className="mb-3">
            <div className="font-sans text-xs uppercase tracking-widest mb-1" style={{ color: C.muted }}>
              Pregunta diagnóstica {path.length + 1}
            </div>
            <div className="font-serif text-lg font-semibold" style={{ color: C.ink }}>{node.q}</div>
            <div className="font-sans text-xs italic mt-2" style={{ color: C.muted }}>{node.note}</div>
          </div>
          <div className="space-y-2">
            {node.options.map((opt, i) => (
              <button key={i} onClick={() => handleChoice(i)}
                style={{
                  width: '100%', textAlign: 'left', padding: 16, borderRadius: 8,
                  background: '#fff', border: `2px solid ${C.border}`,
                  fontFamily: 'Inter, sans-serif', fontSize: 14, color: C.ink,
                  display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                  cursor: 'pointer', transition: 'border-color .15s'
                }}
                onMouseEnter={(e) => e.currentTarget.style.borderColor = C.green}
                onMouseLeave={(e) => e.currentTarget.style.borderColor = C.border}>
                <span>{opt.label}</span>
                <ArrowRight size={16} style={{ color: C.green, flexShrink: 0 }} />
              </button>
            ))}
          </div>
        </div>
      )}

      {leaf && !confirmed && (
        <div className="fade-in mt-4 p-5 rounded-xl"
          style={{ background: C.greenBg, border: `2px solid ${C.green}` }}>
          <div className="font-sans text-xs uppercase tracking-widest mb-1 font-semibold" style={{ color: C.green }}>
            Diseño que tus respuestas implican
          </div>
          <div className="font-serif text-2xl font-bold mb-3" style={{ color: C.ink }}>
            {DESIGNS[leaf].name}
          </div>

          {/* DIAGRAMA SVG */}
          <div style={{ background: '#fff', borderRadius: 8, padding: 12, marginBottom: 10, display: 'flex', justifyContent: 'center', border: `1px solid ${C.border}` }}>
            <DesignDiagram designKey={leaf} />
          </div>

          <div className="font-mono text-sm mb-2 px-3 py-2 rounded inline-block"
            style={{ background: C.gray, color: '#fff' }}>
            {DESIGNS[leaf].eq}
          </div>
          <div className="font-sans text-sm leading-relaxed" style={{ color: C.ink }}>
            {DESIGNS[leaf].note}
          </div>

          <div className="mt-5 p-4 rounded-lg" style={{
            background: '#fff',
            border: `2px solid ${canConfirm ? C.green : C.amber}`
          }}>
            <div style={{
              display: 'flex', alignItems: 'center', justifyContent: 'space-between',
              marginBottom: 8, flexWrap: 'wrap', gap: 8
            }}>
              <div style={{
                fontFamily: 'Inter, sans-serif', fontSize: 13, fontWeight: 700,
                color: C.ink, textTransform: 'uppercase', letterSpacing: 0.5
              }}>
                Justificación obligatoria · mínimo {MIN_JUST} caracteres
              </div>
              <div style={{
                fontFamily: 'JetBrains Mono, monospace', fontSize: 14, fontWeight: 700,
                padding: '4px 10px', borderRadius: 6,
                background: canConfirm ? C.greenBg : C.amberBg,
                color: canConfirm ? C.green : C.amber,
                border: `1px solid ${canConfirm ? C.green : C.amber}`
              }}>
                {justLength} / {MIN_JUST} {canConfirm ? '✓' : ''}
              </div>
            </div>

            <div style={{
              height: 6, borderRadius: 3, background: C.border,
              overflow: 'hidden', marginBottom: 10
            }}>
              <div style={{
                height: '100%',
                width: `${Math.min(100, (justLength / MIN_JUST) * 100)}%`,
                background: canConfirm ? C.green : C.amber,
                transition: 'width .2s, background .2s'
              }} />
            </div>

            <div className="font-sans text-xs mb-2" style={{ color: C.muted, lineHeight: 1.5 }}>
              Argumenta por qué este diseño y no los alternativos, y qué condiciones del escenario lo determinan.
            </div>
            <textarea value={justification}
              onChange={e => setJustification(e.target.value)}
              rows={4}
              placeholder="Ej: El escenario presenta una fuente de variabilidad sistemática que conviene tratar como bloque..."
              style={{
                width: '100%', padding: '12px 16px', borderRadius: 8,
                border: `1px solid ${C.border}`, background: C.cream,
                fontFamily: 'Inter, sans-serif', fontSize: 14, lineHeight: 1.5,
                resize: 'vertical', outline: 'none', color: C.ink
              }} />
          </div>

          {!canConfirm && (
            <div style={{
              marginTop: 10, padding: '10px 14px', borderRadius: 8,
              background: C.amberBg, border: `1px solid ${C.amber}`,
              color: C.amber, fontFamily: 'Inter, sans-serif', fontSize: 13,
              display: 'flex', alignItems: 'center', gap: 8
            }}>
              <AlertCircle size={16} style={{ flexShrink: 0 }} />
              <span>
                <strong>Faltan {MIN_JUST - justLength} caracteres</strong> en la justificación
                para poder confirmar el diseño.
              </span>
            </div>
          )}

          <div className="flex gap-2 mt-3">
            <button onClick={handleReset}
              style={{
                padding: '10px 18px', borderRadius: 8, fontWeight: 600, fontSize: 14,
                background: '#fff', color: C.ink, border: `2px solid ${C.ink}`,
                cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 8,
                fontFamily: 'Inter, sans-serif'
              }}>
              <RotateCcw size={14} /> Reiniciar diagnóstico
            </button>
            <button onClick={handleConfirmClick}
              title={canConfirm ? 'Confirmar este diseño' : `Necesitas ${MIN_JUST - justLength} caracteres más`}
              style={{
                padding: '10px 18px', borderRadius: 8, fontWeight: 700, fontSize: 14,
                background: canConfirm ? C.green : C.disabled,
                color: '#fff', border: 'none',
                cursor: canConfirm ? 'pointer' : 'not-allowed',
                opacity: canConfirm ? 1 : 0.55,
                display: 'flex', alignItems: 'center', gap: 8,
                fontFamily: 'Inter, sans-serif',
                boxShadow: canConfirm ? '0 2px 8px rgba(15,110,86,0.3)' : 'none',
                transition: 'all .2s'
              }}>
              <Check size={14} /> Confirmar diseño
            </button>
          </div>
        </div>
      )}

      {confirmed && (
        <div className="fade-in mt-4 p-5 rounded-xl"
          style={{ background: C.greenBg, border: `2px solid ${C.green}` }}>
          <div className="flex items-center gap-2 mb-2">
            <Check size={20} style={{ color: C.green }} />
            <div className="font-sans text-sm font-semibold" style={{ color: C.green }}>
              Caso registrado
            </div>
          </div>
          <div className="font-serif text-lg font-bold" style={{ color: C.ink }}>
            {DESIGNS[leaf].name}
          </div>
          <div className="font-sans text-xs italic mt-2" style={{ color: C.muted }}>
            {justification.length > 140 ? justification.slice(0, 140) + '...' : justification}
          </div>
          <button onClick={handleReset}
            style={{
              marginTop: 12, background: 'none', border: 'none', cursor: 'pointer',
              color: C.muted, fontSize: 12, textDecoration: 'underline',
              fontFamily: 'Inter, sans-serif'
            }}>
            Modificar respuesta
          </button>
        </div>
      )}
    </div>
  );
}

// =============== CASO AUDITORIA ===============
function AuditCase({ caseData, answer, onAnswer }) {
  const [tags, setTags] = useState(answer?.tags || {});
  const [activeSpan, setActiveSpan] = useState(null);
  const [confirmed, setConfirmed] = useState(answer?.confirmed || false);

  const handleSpanClick = (spanId) => {
    if (confirmed) return;
    setActiveSpan(spanId === activeSpan ? null : spanId);
  };

  const handleAssignTag = (tagKey) => {
    if (!activeSpan) return;
    setTags({ ...tags, [activeSpan]: tagKey });
    setActiveSpan(null);
  };

  const handleRemoveTag = (spanId) => {
    const next = { ...tags };
    delete next[spanId];
    setTags(next);
    setActiveSpan(null);
  };

  const canConfirm = Object.keys(tags).length > 0;

  const handleConfirmClick = () => {
    if (!canConfirm) return;
    setConfirmed(true);
    onAnswer({ tags, confirmed: true });
  };

  return (
    <div className="bg-white rounded-2xl p-7 shadow-md fade-in" style={{ border: `1px solid ${C.coral}` }}>
      <div className="flex items-center gap-2 mb-3 flex-wrap">
        <span className="font-sans text-xs uppercase tracking-widest px-3 py-1 rounded-full"
          style={{ background: C.coral, color: '#fff' }}>
          <Target size={11} style={{ display: 'inline', marginRight: 4 }} /> Parte 2 · Auditoría
        </span>
        <span className="font-sans text-xs italic" style={{ color: C.muted }}>{caseData.area}</span>
      </div>
      <h3 className="text-2xl font-serif font-bold mb-3" style={{ color: C.ink }}>{caseData.title}</h3>

      <div className="p-4 mb-4 rounded-lg font-sans text-sm leading-relaxed"
        style={{ background: C.blueBg, borderLeft: `3px solid ${C.blue}`, color: C.ink }}>
        <div className="font-semibold uppercase text-xs tracking-widest mb-2" style={{ color: C.blue }}>
          Contexto experimental
        </div>
        <div className="space-y-1">
          <div><span className="font-semibold">Objetivo:</span> {caseData.context.objetivo}</div>
          <div><span className="font-semibold">Factores estudiados:</span> {caseData.context.factores}</div>
          <div><span className="font-semibold">Variable respuesta:</span> {caseData.context.respuesta}</div>
          <div><span className="font-semibold">Diseño planeado:</span> {caseData.context.diseno}</div>
        </div>
      </div>

      <div className="p-3 mb-4 rounded-lg font-sans text-sm font-semibold"
        style={{ background: C.coralBg, color: C.coral, border: `1px solid ${C.coral}` }}>
        {caseData.intro}
      </div>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
        <div className="md:col-span-2 p-5 rounded-lg font-serif text-base leading-loose"
          style={{ background: C.cream, border: `1px solid ${C.border}`, color: C.ink }}>
          {caseData.spans.map(span => {
            const tag = tags[span.id];
            const isActive = activeSpan === span.id;
            const isClickable = !confirmed;
            const isMeaningful = span.text.trim().length > 5;
            if (!isMeaningful) return <span key={span.id}>{span.text}</span>;

            let style = {
              cursor: isClickable ? 'pointer' : 'default',
              transition: 'all 0.2s', padding: '2px 4px', borderRadius: 4,
              borderBottom: '2px dotted transparent'
            };
            if (tag) {
              const tagInfo = AUDIT_TAGS[tag];
              style.background = tagInfo.color + '22';
              style.borderBottom = `2px solid ${tagInfo.color}`;
            } else if (isActive) {
              style.background = C.coralBg;
              style.borderBottom = `2px solid ${C.coral}`;
            } else if (isClickable) {
              style.borderBottom = `2px dotted ${C.border}`;
            }

            return (
              <span key={span.id} onClick={() => handleSpanClick(span.id)}
                onMouseEnter={(e) => { if (isClickable && !tag) e.currentTarget.style.background = C.coralBg; }}
                onMouseLeave={(e) => { if (isClickable && !tag) e.currentTarget.style.background = 'transparent'; }}
                style={style}>
                {span.text}
                {tag && (
                  <span className="font-sans text-xs ml-1 px-1.5 py-0.5 rounded font-semibold"
                    style={{ background: AUDIT_TAGS[tag].color, color: '#fff' }}>
                    {AUDIT_TAGS[tag].label}
                  </span>
                )}
              </span>
            );
          })}
        </div>

        <div className="space-y-2">
          <div className="font-sans text-xs uppercase tracking-widest mb-1 font-semibold" style={{ color: C.muted }}>
            <Tag size={11} style={{ display: 'inline', marginRight: 4 }} />
            {activeSpan ? 'Asigna categoría' : 'Categorías de error'}
          </div>
          {activeSpan && tags[activeSpan] && (
            <button onClick={() => handleRemoveTag(activeSpan)}
              style={{
                width: '100%', textAlign: 'left', padding: '8px 12px', borderRadius: 8,
                background: '#fff', color: C.coral, border: `1px dashed ${C.coral}`,
                fontFamily: 'Inter, sans-serif', fontSize: 12, cursor: 'pointer',
                display: 'flex', alignItems: 'center', gap: 8
              }}>
              <X size={12} /> Quitar etiqueta
            </button>
          )}
          {Object.entries(AUDIT_TAGS).map(([key, tag]) => {
            const enabled = activeSpan && !confirmed;
            return (
              <button key={key} onClick={() => { if (enabled) handleAssignTag(key); }}
                style={{
                  width: '100%', textAlign: 'left', padding: '8px 12px', borderRadius: 8,
                  background: '#fff', border: `1px solid ${C.border}`, borderLeft: `4px solid ${tag.color}`,
                  fontFamily: 'Inter, sans-serif', fontSize: 12, fontWeight: 600, color: C.ink,
                  cursor: enabled ? 'pointer' : 'not-allowed',
                  opacity: enabled ? 1 : 0.4
                }}>
                {tag.label}
              </button>
            );
          })}
          {!activeSpan && !confirmed && (
            <div className="font-sans text-xs italic mt-3" style={{ color: C.muted }}>
              Haz clic en una frase del reporte para etiquetarla.
            </div>
          )}
        </div>
      </div>

      <div className="mt-5 flex items-center justify-between flex-wrap gap-3">
        <div className="font-sans text-xs" style={{ color: C.muted }}>
          {Object.keys(tags).length} frase{Object.keys(tags).length === 1 ? '' : 's'} etiquetada{Object.keys(tags).length === 1 ? '' : 's'}
        </div>
        {!confirmed ? (
          <button onClick={handleConfirmClick}
            style={{
              padding: '8px 16px', borderRadius: 8, fontWeight: 600, fontSize: 14,
              background: canConfirm ? C.coral : C.disabled, color: '#fff', border: 'none',
              cursor: canConfirm ? 'pointer' : 'not-allowed',
              opacity: canConfirm ? 1 : 0.65,
              display: 'flex', alignItems: 'center', gap: 8,
              fontFamily: 'Inter, sans-serif'
            }}>
            <Check size={14} /> Confirmar análisis
          </button>
        ) : (
          <div className="flex items-center gap-3">
            <span className="font-sans text-sm font-semibold flex items-center gap-1" style={{ color: C.green }}>
              <Check size={16} /> Análisis registrado
            </span>
            <button onClick={() => setConfirmed(false)}
              style={{
                background: 'none', border: 'none', cursor: 'pointer',
                color: C.muted, fontSize: 12, textDecoration: 'underline',
                fontFamily: 'Inter, sans-serif'
              }}>
              Modificar
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

// =============== PANTALLA DE REVISION ===============
function ReviewScreen({ student, detectiveAnswers, auditAnswers, onEdit, onSubmit, onBackToWorkshop, submitting }) {
  return (
    <div className="fade-in">
      <div className="mb-6">
        <div className="inline-block px-4 py-1 mb-3 rounded-full font-sans text-xs uppercase tracking-widest"
          style={{ background: C.amber, color: '#fff' }}>
          <FileCheck size={12} style={{ display: 'inline', marginRight: 4 }} /> Revisión final
        </div>
        <h2 className="text-3xl font-serif font-bold mb-2" style={{ color: C.ink }}>
          Audita tus respuestas antes de entregar
        </h2>
        <p className="font-sans text-sm" style={{ color: C.muted }}>
          {student.name} · cédula {student.cedula}. Haz clic en cualquier caso para volver y modificarlo.
          La entrega definitiva es irreversible.
        </p>
      </div>

      <div className="bg-white rounded-2xl p-5 mb-5" style={{ border: `1px solid ${C.green}` }}>
        <div className="font-sans text-xs uppercase tracking-widest font-semibold mb-3" style={{ color: C.green }}>
          Parte 1 · Selección de diseño
        </div>
        <div className="space-y-2">
          {CASES.map((c, i) => {
            const ans = detectiveAnswers[c.id];
            return (
              <button key={c.id} onClick={() => onEdit('detective', i)}
                style={{
                  width: '100%', textAlign: 'left', padding: '12px 14px', borderRadius: 8,
                  background: ans?.confirmed ? C.greenBg : '#fff8eb',
                  border: `1px solid ${ans?.confirmed ? C.green : C.amber}`,
                  cursor: 'pointer', display: 'flex', alignItems: 'center',
                  gap: 12, fontFamily: 'Inter, sans-serif', transition: 'transform .15s'
                }}
                onMouseEnter={e => e.currentTarget.style.transform = 'translateX(4px)'}
                onMouseLeave={e => e.currentTarget.style.transform = 'translateX(0)'}>
                <div style={{ flexShrink: 0 }}>
                  {ans?.confirmed
                    ? <Check size={18} style={{ color: C.green }} />
                    : <AlertCircle size={18} style={{ color: C.amber }} />}
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.ink }}>{c.title}</div>
                  <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>
                    {ans?.confirmed ? `Diseño: ${DESIGNS[ans.leaf]?.name}` : 'Sin confirmar'}
                  </div>
                </div>
                <Edit3 size={14} style={{ color: C.muted, flexShrink: 0 }} />
              </button>
            );
          })}
        </div>
      </div>

      <div className="bg-white rounded-2xl p-5 mb-6" style={{ border: `1px solid ${C.coral}` }}>
        <div className="font-sans text-xs uppercase tracking-widest font-semibold mb-3" style={{ color: C.coral }}>
          Parte 2 · Auditoría
        </div>
        <div className="space-y-2">
          {AUDIT_CASES.map((c, i) => {
            const ans = auditAnswers[c.id];
            const tagCount = ans?.tags ? Object.keys(ans.tags).length : 0;
            return (
              <button key={c.id} onClick={() => onEdit('audit', i)}
                style={{
                  width: '100%', textAlign: 'left', padding: '12px 14px', borderRadius: 8,
                  background: ans?.confirmed ? C.coralBg : '#fff8eb',
                  border: `1px solid ${ans?.confirmed ? C.coral : C.amber}`,
                  cursor: 'pointer', display: 'flex', alignItems: 'center',
                  gap: 12, fontFamily: 'Inter, sans-serif', transition: 'transform .15s'
                }}
                onMouseEnter={e => e.currentTarget.style.transform = 'translateX(4px)'}
                onMouseLeave={e => e.currentTarget.style.transform = 'translateX(0)'}>
                <div style={{ flexShrink: 0 }}>
                  {ans?.confirmed
                    ? <Check size={18} style={{ color: C.coral }} />
                    : <AlertCircle size={18} style={{ color: C.amber }} />}
                </div>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13, fontWeight: 600, color: C.ink }}>{c.title}</div>
                  <div style={{ fontSize: 12, color: C.muted, marginTop: 2 }}>
                    {ans?.confirmed ? `${tagCount} ${tagCount === 1 ? 'frase etiquetada' : 'frases etiquetadas'}` : 'Sin confirmar'}
                  </div>
                </div>
                <Edit3 size={14} style={{ color: C.muted, flexShrink: 0 }} />
              </button>
            );
          })}
        </div>
      </div>

      <div className="p-4 mb-5 rounded-lg font-sans text-sm"
        style={{ background: C.amberBg, border: `1px solid ${C.amber}`, color: C.amber }}>
        <div className="flex items-start gap-2">
          <AlertCircle size={18} style={{ flexShrink: 0, marginTop: 2 }} />
          <div>
            <strong>Antes de entregar:</strong> revisa que cada caso esté confirmado y que tus
            justificaciones expresen lo que realmente quisiste argumentar. Una vez entregado el
            taller, las respuestas quedan registradas para calificación.
          </div>
        </div>
      </div>

      <div className="flex justify-between gap-3">
        <button onClick={onBackToWorkshop}
          style={{
            padding: '12px 24px', borderRadius: 8, fontWeight: 600, fontSize: 14,
            background: '#fff', color: C.ink, border: `2px solid ${C.ink}`,
            cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 8,
            fontFamily: 'Inter, sans-serif'
          }}>
          <ArrowLeft size={16} /> Volver a editar
        </button>
        <button onClick={() => { if (!submitting) onSubmit(); }}
          style={{
            padding: '12px 24px', borderRadius: 8, fontWeight: 600, fontSize: 14,
            background: C.green, color: '#fff', border: 'none',
            cursor: submitting ? 'not-allowed' : 'pointer',
            opacity: submitting ? 0.6 : 1,
            display: 'flex', alignItems: 'center', gap: 8,
            fontFamily: 'Inter, sans-serif'
          }}>
          {submitting ? 'Enviando...' : 'Entregar definitivamente'} <ClipboardCheck size={16} />
        </button>
      </div>
    </div>
  );
}

// =============== PANTALLA FINAL ===============
function DoneScreen({ student, score, onExit }) {
  const pct = (score.auto / score.totalMax) * 100;
  const grade = score.scaled >= 4.5 ? 'A' : score.scaled >= 3.5 ? 'B' : score.scaled >= 3 ? 'C' : score.scaled >= 2 ? 'D' : 'E';
  const gradeColor = { A: C.green, B: C.blue, C: C.amber, D: C.coral, E: '#7f0000' }[grade];
  const gradeBg = { A: C.greenBg, B: C.blueBg, C: C.amberBg, D: C.coralBg, E: '#FCEBEB' }[grade];

  return (
    <div className="max-w-2xl mx-auto py-12 text-center fade-in">
      <div className="inline-flex items-center justify-center w-20 h-20 rounded-full mb-6" style={{ background: C.green }}>
        <Check size={40} style={{ color: '#fff' }} />
      </div>
      <h2 className="text-4xl font-serif font-bold mb-2" style={{ color: C.ink }}>
        Entrega registrada
      </h2>
      <p className="font-sans mb-8" style={{ color: C.muted }}>
        Gracias {student.name}. Tu taller fue guardado bajo la cédula {student.cedula}.
      </p>
      <div className="rounded-2xl p-8 shadow-md mb-6" style={{ background: gradeBg, border: `2px solid ${gradeColor}` }}>
        <div className="font-sans text-xs uppercase tracking-widest mb-2 font-semibold" style={{ color: gradeColor }}>
          Calificación automática
        </div>
        <div className="flex items-baseline justify-center gap-3 mb-4">
          <span className="font-serif font-bold" style={{ color: gradeColor, fontSize: '4rem', lineHeight: 1 }}>
            {score.scaled.toFixed(2)}
          </span>
          <span className="text-2xl" style={{ color: C.muted }}>/ 5.0</span>
          <span className="font-serif font-bold ml-4" style={{ color: gradeColor, fontSize: '3rem', lineHeight: 1 }}>{grade}</span>
        </div>
        <div className="font-sans text-sm mb-4" style={{ color: gradeColor }}>
          {score.auto.toFixed(1)} / {score.totalMax} puntos automáticos. Las justificaciones serán revisadas por el docente.
        </div>
        <div className="grid grid-cols-2 gap-3 mb-3 text-sm font-sans">
          <div className="p-3 rounded" style={{ background: 'rgba(255,255,255,0.6)' }}>
            <div className="font-semibold" style={{ color: C.ink }}>Detective</div>
            <div className="font-bold" style={{ color: C.green }}>{score.detPoints.toFixed(1)} / {CASES.length}</div>
          </div>
          <div className="p-3 rounded" style={{ background: 'rgba(255,255,255,0.6)' }}>
            <div className="font-semibold" style={{ color: C.ink }}>Auditoría</div>
            <div className="font-bold" style={{ color: C.coral }}>{score.audPoints.toFixed(1)} / {AUDIT_CASES.length}</div>
          </div>
        </div>
        <div className="h-3 rounded-full overflow-hidden" style={{ background: 'rgba(255,255,255,0.6)' }}>
          <div className="h-full transition-all" style={{ background: gradeColor, width: `${pct}%` }} />
        </div>
      </div>
      <button onClick={onExit}
        style={{
          padding: '12px 24px', borderRadius: 8, fontWeight: 600, fontSize: 14,
          background: C.green, color: '#fff', border: 'none', cursor: 'pointer',
          fontFamily: 'Inter, sans-serif'
        }}>
        Salir
      </button>
    </div>
  );
}

// =============== FLUJO DOCENTE ===============
function TeacherFlow({ onExit }) {
  const [authed, setAuthed] = useState(false);
  const [pwd, setPwd] = useState('');
  const [error, setError] = useState('');
  const [submissions, setSubmissions] = useState([]);
  const [loading, setLoading] = useState(false);

  const loadSubmissions = async () => {
    setLoading(true);
    try {
      const res = await window.storage.list(STORAGE_PREFIX, true);
      const keys = res?.keys || [];
      const subs = [];
      for (const k of keys) {
        try {
          const item = await window.storage.get(k, true);
          if (item?.value) subs.push(JSON.parse(item.value));
        } catch (e) { /* skip */ }
      }
      subs.sort((a, b) => new Date(b.submittedAt) - new Date(a.submittedAt));
      setSubmissions(subs);
    } catch (e) { setError('Error: ' + e.message); }
    setLoading(false);
  };

  useEffect(() => { if (authed) loadSubmissions(); }, [authed]);

  const handleAuth = () => {
    if (pwd === TEACHER_PASSWORD) { setAuthed(true); setError(''); }
    else setError('Contraseña incorrecta');
  };

  const exportExcel = () => {
    if (submissions.length === 0) return;

    const summary = submissions.map(s => ({
      'Cédula': s.cedula, 'Nombre': s.name, 'Programa': s.program || '',
      'Fecha entrega': new Date(s.submittedAt).toLocaleString('es-CO'),
      [`Parte 1 (sobre ${CASES.length})`]: s.score.detPoints.toFixed(2),
      [`Parte 2 (sobre ${AUDIT_CASES.length})`]: s.score.audPoints.toFixed(2),
      [`Total auto (sobre ${CASES.length + AUDIT_CASES.length})`]: s.score.auto.toFixed(2),
      'Nota sobre 5.0': s.score.scaled.toFixed(2),
      'Ajuste docente (-0.5 a +0.5)': '',
      'Nota final': ''
    }));

    const detective = submissions.map(s => {
      const row = { 'Cédula': s.cedula, 'Nombre': s.name };
      CASES.forEach(c => {
        const d = s.score.detDetail?.[c.id];
        row[`${c.id}_diseño`] = d?.leaf ? DESIGNS[d.leaf]?.name : '(sin respuesta)';
        row[`${c.id}_correcto`] = DESIGNS[c.expected]?.name;
        row[`${c.id}_pts`] = d?.points?.toFixed(2) || '0.00';
      });
      return row;
    });

    const justifications = [];
    submissions.forEach(s => {
      CASES.forEach(c => {
        const d = s.score.detDetail?.[c.id];
        if (d?.justification) {
          justifications.push({
            'Cédula': s.cedula, 'Nombre': s.name, 'Caso': c.id,
            'Título': c.title,
            'Diseño elegido': DESIGNS[d.leaf]?.name || 'Sin respuesta',
            'Diseño correcto': DESIGNS[c.expected]?.name,
            'Acertó': d.leaf === c.expected ? 'SÍ' : 'NO',
            'Justificación': d.justification
          });
        }
      });
    });

    const audit = submissions.map(s => {
      const row = { 'Cédula': s.cedula, 'Nombre': s.name };
      AUDIT_CASES.forEach(c => {
        const d = s.score.audDetail?.[c.id];
        row[`${c.id}_correctas`] = d?.correctTags ?? 0;
        row[`${c.id}_erróneas`] = d?.wrongTags ?? 0;
        row[`${c.id}_total`] = d?.totalErrors ?? 0;
        row[`${c.id}_pts`] = d?.points?.toFixed(2) || '0.00';
      });
      return row;
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summary), 'Resumen');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(detective), 'Parte 1 detalle');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(audit), 'Parte 2 detalle');
    XLSX.utils.book_append_sheet(wb,
      XLSX.utils.json_to_sheet(justifications.length ? justifications : [{ Mensaje: 'Sin justificaciones registradas' }]),
      'Justificaciones'
    );
    XLSX.writeFile(wb, `Taller_DOE_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const deleteOne = async (cedula) => {
    if (!confirm(`¿Eliminar entrega ${cedula}?`)) return;
    await window.storage.delete(`${STORAGE_PREFIX}${cedula}`, true);
    loadSubmissions();
  };

  if (!authed) {
    return (
      <div className="max-w-md mx-auto py-12 fade-in">
        <button onClick={onExit} style={{
          background: 'none', border: 'none', cursor: 'pointer',
          color: C.muted, fontSize: 14, marginBottom: 24,
          display: 'flex', alignItems: 'center', gap: 4, fontFamily: 'Inter, sans-serif'
        }}>
          <ArrowLeft size={14} /> Volver
        </button>
        <div className="bg-white rounded-2xl p-8 shadow-lg" style={{ border: `1px solid ${C.gray}` }}>
          <ShieldAlert size={36} style={{ color: C.gray }} className="mb-3" />
          <h2 className="text-2xl font-serif font-bold mb-1" style={{ color: C.ink }}>Acceso docente</h2>
          <p className="font-sans text-sm mb-5" style={{ color: C.muted }}>
            Ingresa la contraseña para ver entregas y exportar a Excel.
          </p>
          <Field label="Contraseña" value={pwd} onChange={setPwd} type="password" placeholder="••••••••" />
          {error && <div className="font-sans text-sm mt-2" style={{ color: C.coral }}>{error}</div>}
          <button onClick={handleAuth}
            style={{
              width: '100%', marginTop: 16, padding: '12px 24px', borderRadius: 8,
              background: C.gray, color: '#fff', border: 'none', cursor: 'pointer',
              fontFamily: 'Inter, sans-serif', fontWeight: 600, fontSize: 14
            }}>
            Ingresar
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="fade-in">
      <div className="flex justify-between items-center mb-6">
        <button onClick={onExit} style={{
          background: 'none', border: 'none', cursor: 'pointer',
          color: C.muted, fontSize: 14, display: 'flex', alignItems: 'center', gap: 4,
          fontFamily: 'Inter, sans-serif'
        }}>
          <ArrowLeft size={14} /> Volver
        </button>
        <div className="flex gap-2">
          <button onClick={loadSubmissions}
            style={{
              padding: '8px 16px', borderRadius: 8, fontWeight: 600, fontSize: 14,
              background: '#fff', color: C.ink, border: `2px solid ${C.ink}`, cursor: 'pointer',
              display: 'flex', alignItems: 'center', gap: 8, fontFamily: 'Inter, sans-serif'
            }}>
            <RefreshCw size={14} /> Refrescar
          </button>
          <button onClick={() => { if (submissions.length > 0) exportExcel(); }}
            style={{
              padding: '8px 16px', borderRadius: 8, fontWeight: 600, fontSize: 14,
              background: submissions.length === 0 ? C.disabled : C.green,
              color: '#fff', border: 'none',
              cursor: submissions.length === 0 ? 'not-allowed' : 'pointer',
              opacity: submissions.length === 0 ? 0.5 : 1,
              display: 'flex', alignItems: 'center', gap: 8, fontFamily: 'Inter, sans-serif'
            }}>
            <Download size={14} /> Descargar Excel
          </button>
        </div>
      </div>

      <h2 className="text-3xl font-serif font-bold mb-1" style={{ color: C.ink }}>Panel docente</h2>
      <p className="font-sans mb-6" style={{ color: C.muted }}>
        {submissions.length} entrega{submissions.length === 1 ? '' : 's'}
      </p>

      <div className="grid grid-cols-3 gap-4 mb-6">
        <StatCard label="Entregas" value={submissions.length} icon={Users} color={C.gray} />
        <StatCard label="Promedio /5.0" value={
          submissions.length > 0
            ? (submissions.reduce((acc, s) => acc + s.score.scaled, 0) / submissions.length).toFixed(2)
            : '0.00'
        } icon={Award} color={C.blue} />
        <StatCard label="Aprobados ≥3.0" value={
          submissions.length > 0
            ? `${Math.round((submissions.filter(s => s.score.scaled >= 3).length / submissions.length) * 100)}%`
            : '0%'
        } icon={ClipboardCheck} color={C.green} />
      </div>

      {loading ? (
        <div className="text-center py-12 font-sans" style={{ color: C.muted }}>Cargando...</div>
      ) : submissions.length === 0 ? (
        <div className="bg-white rounded-2xl p-12 text-center" style={{ border: `1px dashed ${C.border}` }}>
          <p className="font-sans" style={{ color: C.muted }}>
            Aún no hay entregas. Comparte el enlace del taller con tus estudiantes.
          </p>
        </div>
      ) : (
        <div className="bg-white rounded-2xl overflow-hidden" style={{ border: `1px solid ${C.border}` }}>
          <table className="w-full">
            <thead style={{ background: C.gray }}>
              <tr>
                {['Cédula', 'Nombre', `P1/${CASES.length}`, `P2/${AUDIT_CASES.length}`, `Total/${CASES.length + AUDIT_CASES.length}`, 'Nota/5', 'Fecha', ''].map(h => (
                  <th key={h} className="px-3 py-3 text-left font-sans text-xs uppercase tracking-widest" style={{ color: '#fff' }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {submissions.map((s, i) => (
                <tr key={s.cedula} style={{
                  background: i % 2 === 0 ? '#fff' : C.cream, borderTop: `1px solid ${C.border}`
                }}>
                  <td className="px-3 py-3 font-mono text-sm" style={{ color: C.ink }}>{s.cedula}</td>
                  <td className="px-3 py-3 font-sans text-sm" style={{ color: C.ink }}>{s.name}</td>
                  <td className="px-3 py-3 font-mono text-sm" style={{ color: C.muted }}>{s.score.detPoints.toFixed(1)}</td>
                  <td className="px-3 py-3 font-mono text-sm" style={{ color: C.muted }}>{s.score.audPoints.toFixed(1)}</td>
                  <td className="px-3 py-3 font-mono text-sm font-bold" style={{ color: C.green }}>{s.score.auto.toFixed(1)}</td>
                  <td className="px-3 py-3 font-mono text-sm font-bold" style={{
                    color: s.score.scaled >= 3 ? C.green : C.coral
                  }}>{s.score.scaled.toFixed(2)}</td>
                  <td className="px-3 py-3 font-sans text-xs" style={{ color: C.muted }}>
                    {new Date(s.submittedAt).toLocaleString('es-CO', {
                      day: '2-digit', month: 'short', hour: '2-digit', minute: '2-digit'
                    })}
                  </td>
                  <td className="px-3 py-3">
                    <button onClick={() => deleteOne(s.cedula)}
                      style={{ background: 'none', border: 'none', cursor: 'pointer', padding: 4 }}>
                      <Trash2 size={14} style={{ color: C.coral }} />
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      <div className="mt-6 p-5 rounded-xl font-sans text-sm"
        style={{ background: C.amberBg, border: `1px solid ${C.amber}`, color: C.amber }}>
        <div className="font-semibold mb-2">Estructura del Excel descargable</div>
        <div style={{ color: C.ink }}>
          Hoja 1 (Resumen): nota lista con columnas vacías para tu ajuste manual.
          Hoja 2 (Parte 1 detalle): diseño elegido y diseño correcto por caso.
          Hoja 3 (Parte 2 detalle): conteo de etiquetas correctas y erróneas.
          Hoja 4 (Justificaciones): cada justificación lista para revisar y bonificar.
        </div>
      </div>
    </div>
  );
}

function StatCard({ label, value, icon: Icon, color }) {
  return (
    <div className="bg-white rounded-xl p-4" style={{ border: `1px solid ${color}`, borderLeft: `4px solid ${color}` }}>
      <Icon size={18} style={{ color }} className="mb-1" />
      <div className="font-sans text-xs uppercase tracking-widest" style={{ color: C.muted }}>{label}</div>
      <div className="font-serif text-2xl font-bold" style={{ color: C.ink }}>{value}</div>
    </div>
  );
}
