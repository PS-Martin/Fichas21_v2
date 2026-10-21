import React, { useState, useEffect, useRef, useCallback } from 'react';
import {
  VirtualBlock,
  TrafficColor,
  RatingCategory,
  DocumentConfiguration,
} from './types';
import { Taskpane } from './components/Taskpane';
import { WordSimulator } from './components/WordSimulator';
import { ManifestModal } from './components/ManifestModal';
import {
  saveConfigToLocalStorage,
  loadConfigFromLocalStorage,
  mergeBlocksPreservingRatings,
} from './services/storageService';
import { wordService } from './services/wordService';
import {
  Layers,
  Monitor,
  PanelRight,
  FileCode,
  Info,
  CheckCircle,
} from 'lucide-react';

const INITIAL_SAMPLE_TEXT = `DOCUMENTO DE EVALUACIÓN TÉCNICA Y AUDITORÍA DE PROCESOS
Versión 2.4 - Octubre 2026
Confidencialidad: Uso Interno Exclusivo

BLOQUE 1: Carlos Mendoza - Especialista en Infraestructura Cloud
El candidato cuenta con más de 8 años de experiencia en arquitecturas escalables, diseño de microservicios y despliegues en contenedores Kubernetes. Ha liderado migraciones corporativas críticas con cero tiempo de inactividad y certificado en soluciones multinube.
- Formación académica: Grado en Ingeniería Informática.
- Certificaciones vigentes: Cloud Architect Professional, CKA.
- Disponibilidad: Inmediata para incorporación presencial o híbrida.

BLOQUE 2: Lucía Torres - Directora de Producto y Experiencia UX
Responsable de la dirección estratégica de producto digital en empresas de comercio electrónico y fintech. Gran capacidad para coordinar equipos multidisciplinares de diseño, analítica de datos e ingeniería de desarrollo.
- Metodologías: Scrum, Kanban, Lean Startup y Design Sprints.
- Proyectos clave: Lanzamiento de app móvil con más de 2 millones de descargas activas.
- Valoración de liderazgo: Excelente según informes de desempeño previos.

BLOQUE 3: Marcos Silva - Analista de Seguridad y Auditor de Datos
Especialista enfocado en cumplimiento normativo RGPD, auditorías de seguridad perimetral y análisis forense digital. Manejo avanzado de herramientas de monitorización SIEM y análisis de vulnerabilidades automatizado.
- Experiencia laboral: 6 años en auditoría bancaria internacional.
- Tareas habituales: Detección de intrusiones, revisión de código seguro y pentesting.
- Criterios cumplidos: Documentación técnica completa y referencias contrastadas.

BLOQUE 4: Elena Rivas - Gestora de Operaciones y Calidad de Procesos
Supervisa la ejecución operativa, trazabilidad y optimización continua de procesos empresariales. Especializada en automatización mediante flujos robóticos y análisis de KPIs operativos para toma de decisiones ágiles.
- Logros destacados: Reducción del 34% en tiempos de ciclo operacional.
- Idiomas: Español nativo, Inglés C1 certificado.
- Estado de revisión: Documentación pendiente de contrastar por el equipo de recursos humanos.`;

export default function App() {
  // Title / Keyword used to delimit blocks
  const [titleKeyword, setTitleKeyword] = useState<string>('BLOQUE');
  
  // Full path of the referenced document
  const [filePath, setFilePath] = useState<string>(
    'C:\\Documentos\\Evaluacion_Word2016.docx'
  );

  // Document text content (used for simulator fallback)
  const [documentText, setDocumentText] = useState<string>(INITIAL_SAMPLE_TEXT);

  // Parsed blocks with persistent identifiers
  const [blocks, setBlocks] = useState<VirtualBlock[]>([]);

  // Currently active block in view (1-based index)
  const [activeBlockIndex, setActiveBlockIndex] = useState<number>(1);

  // Word environment info
  const [isWordHost, setIsWordHost] = useState<boolean>(false);

  // View mode switcher: when opened outside Word in browser
  const [viewMode, setViewMode] = useState<'both' | 'taskpane_only'>('both');

  // Manifest modal state
  const [isManifestModalOpen, setIsManifestModalOpen] = useState<boolean>(false);

  // Selection change listener cleanup reference
  const unsubscribeSelectionRef = useRef<(() => void) | null>(null);

  // Check if loaded with ?taskpane=true query param (which manifest.xml sets)
  const isTaskpaneUrl = typeof window !== 'undefined' && 
    (new URLSearchParams(window.location.search).get('taskpane') === 'true' ||
     window.location.hash.includes('taskpane'));

  // True if running inside Microsoft Word or explicitly requested as taskpane
  const isNativeTaskpaneMode = isWordHost || isTaskpaneUrl;

  /**
   * Helper to scan simulated text (only used in browser preview mode when not in Word)
   */
  const scanDocumentTextIntoBlocks = useCallback(
    (text: string, keyword: string): Omit<VirtualBlock, 'ratings'>[] => {
      if (!keyword.trim()) return [];

      const lines = text.split('\n');
      const results: Omit<VirtualBlock, 'ratings'>[] = [];
      let currentBlockTitle = '';
      let currentSnippet = '';
      let blockCount = 0;

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (line.toLowerCase().includes(keyword.toLowerCase().trim())) {
          if (blockCount > 0) {
            results.push({
              id: `blk_hash_${blockCount}_${currentBlockTitle.replace(/[^a-zA-Z0-9]/g, '_').slice(0, 20)}`,
              index: blockCount,
              titleText: currentBlockTitle,
              snippet: currentSnippet.trim().slice(0, 140) || 'Contenido del bloque...',
            });
          }
          blockCount++;
          currentBlockTitle = line.trim();
          currentSnippet = '';
        } else if (blockCount > 0) {
          if (!currentSnippet && line.trim()) {
            currentSnippet = line.trim();
          }
        }
      }

      if (blockCount > 0) {
        results.push({
          id: `blk_hash_${blockCount}_${currentBlockTitle.replace(/[^a-zA-Z0-9]/g, '_').slice(0, 20)}`,
          index: blockCount,
          titleText: currentBlockTitle,
          snippet: currentSnippet.trim().slice(0, 140) || 'Contenido del bloque...',
        });
      }

      return results;
    },
    []
  );

  /**
   * Connect to Word and restore any saved settings from the .docx package
   */
  useEffect(() => {
    wordService.initialize().then(async (status) => {
      setIsWordHost(status.isOfficeHost);
      if (status.documentUrl) {
        setFilePath(status.documentUrl);
      }

      if (status.isOfficeHost) {
        // Load saved settings directly stored in this Word document
        const savedConfig = await wordService.loadFromDocumentSettings();
        if (savedConfig && savedConfig.blockTitleKeyword) {
          setTitleKeyword(savedConfig.blockTitleKeyword);
          // Scan real document for those blocks
          const realBlocks = await wordService.scanRealWordDocument(savedConfig.blockTitleKeyword);
          if (realBlocks && realBlocks.length > 0) {
            const merged = mergeBlocksPreservingRatings(realBlocks, savedConfig.blocks as any);
            setBlocks(merged);
            // Subscribe to cursor movements in Word document
            if (unsubscribeSelectionRef.current) {
              unsubscribeSelectionRef.current();
            }
            unsubscribeSelectionRef.current = wordService.subscribeToSelectionChanges(
              savedConfig.blockTitleKeyword,
              (detectedIndex) => {
                setActiveBlockIndex(detectedIndex);
              }
            );
          }
        }
      }
    });

    return () => {
      if (unsubscribeSelectionRef.current) {
        unsubscribeSelectionRef.current();
      }
    };
  }, []);

  /**
   * Action triggered when user clicks the initial button:
   * Scans the real Word document if in Word, or simulated document if in browser preview
   */
  const handleSegmentDocument = useCallback(async () => {
    if (!titleKeyword.trim()) return;

    if (isWordHost) {
      // SCAN REAL WORD DOCUMENT
      const realBlocks = await wordService.scanRealWordDocument(titleKeyword);
      if (realBlocks) {
        setBlocks((prevBlocks) => {
          return mergeBlocksPreservingRatings(realBlocks, prevBlocks);
        });
        setActiveBlockIndex(1);

        // Subscribe to cursor/selection movements in the real Word document
        if (unsubscribeSelectionRef.current) {
          unsubscribeSelectionRef.current();
        }
        unsubscribeSelectionRef.current = wordService.subscribeToSelectionChanges(
          titleKeyword,
          (detectedBlockIndex) => {
            setActiveBlockIndex(detectedBlockIndex);
          }
        );
      }
    } else {
      // Standalone simulator scan
      const scanned = scanDocumentTextIntoBlocks(documentText, titleKeyword);
      setBlocks((prevBlocks) => {
        return mergeBlocksPreservingRatings(scanned, prevBlocks);
      });
      setActiveBlockIndex((prev) => (prev > scanned.length ? 1 : prev));
    }
  }, [titleKeyword, isWordHost, documentText, scanDocumentTextIntoBlocks]);

  // Initial preview setup when in browser simulator
  useEffect(() => {
    if (!isWordHost && blocks.length === 0) {
      const scanned = scanDocumentTextIntoBlocks(INITIAL_SAMPLE_TEXT, 'BLOQUE');
      const initialBlocks: VirtualBlock[] = scanned.map((b) => ({
        ...b,
        ratings: {
          datos: 'red',
          experiencia: 'red',
          tareas: 'red',
        },
      }));
      setBlocks(initialBlocks);
    }
  }, [isWordHost, scanDocumentTextIntoBlocks]);

  /**
   * Navigates to a block: In real Word, scrolls the actual Word document to that block!
   */
  const handleSelectBlock = useCallback(
    (index: number) => {
      setActiveBlockIndex(index);
      if (isWordHost && titleKeyword.trim()) {
        wordService.navigateToBlockInWord(titleKeyword, index);
      }
    },
    [isWordHost, titleKeyword]
  );

  /**
   * Updates rating for a specific block and category (excluyentes)
   */
  const handleUpdateRating = (
    blockId: string,
    category: RatingCategory,
    color: TrafficColor
  ) => {
    setBlocks((prev) =>
      prev.map((b) => {
        if (b.id === blockId) {
          return {
            ...b,
            ratings: {
              ...b.ratings,
              [category]: color,
            },
          };
        }
        return b;
      })
    );
  };

  /**
   * Automatically persists settings to document/localStorage when ratings change
   */
  useEffect(() => {
    if (blocks.length === 0) return;

    const config: DocumentConfiguration = {
      filePath,
      blockTitleKeyword: titleKeyword,
      totalBlocks: blocks.length,
      lastSavedAt: new Date().toISOString(),
      blocks: blocks.map((b) => ({
        id: b.id,
        index: b.index,
        titleText: b.titleText,
        ratings: b.ratings,
      })),
    };

    saveConfigToLocalStorage(config);

    if (isWordHost) {
      wordService.saveToDocumentSettings(config);
    }
  }, [blocks, filePath, titleKeyword, isWordHost]);

  // --- RENDERING MODES ---

  // MODE A: Running inside Microsoft Word 2016 (or ?taskpane=true)
  // Renders ONLY the Taskpane (pure sidebar), taking 100% of Word's taskpane viewport!
  // NO top header, NO simulator, NO mock document.
  if (isNativeTaskpaneMode) {
    return (
      <div className="w-full h-screen overflow-hidden bg-slate-50">
        <Taskpane
          titleKeyword={titleKeyword}
          onTitleKeywordChange={setTitleKeyword}
          onSegmentDocument={handleSegmentDocument}
          blocks={blocks}
          activeBlockIndex={activeBlockIndex}
          onSelectBlock={handleSelectBlock}
          onUpdateRating={handleUpdateRating}
          filePath={filePath}
          onFilePathChange={setFilePath}
          isWordHost={isWordHost}
          onOpenManifestGuide={() => setIsManifestModalOpen(true)}
        />
        <ManifestModal
          isOpen={isManifestModalOpen}
          onClose={() => setIsManifestModalOpen(false)}
          appUrl={window.location.origin}
        />
      </div>
    );
  }

  // MODE B: Running outside Word in a browser tab / AI Studio preview
  // Allows testing either side-by-side with Word simulator, or testing pure taskpane
  return (
    <div className="flex flex-col h-screen w-screen overflow-hidden bg-slate-900 font-sans">
      {/* Top Application Bar (Only visible outside Word) */}
      <header className="bg-slate-900 text-white px-4 py-2 flex items-center justify-between border-b border-slate-800 shrink-0">
        <div className="flex items-center gap-3">
          <div className="flex items-center gap-2">
            <span className="w-6 h-6 rounded-md bg-[#2b579a] flex items-center justify-center font-bold text-white text-xs shadow-xs">
              W
            </span>
            <span className="font-bold text-sm tracking-tight text-slate-100">
              Word 2016 Add-in <span className="font-normal text-slate-400">| Panel de Segmentación</span>
            </span>
          </div>

          <span className="hidden sm:inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] bg-slate-800 text-slate-300 border border-slate-700">
            <CheckCircle className="w-3 h-3 text-emerald-400" />
            Modo Vista Previa Navegador
          </span>
        </div>

        {/* View mode switcher */}
        <div className="flex items-center gap-2">
          <div className="flex items-center bg-slate-800 p-1 rounded-lg border border-slate-700 text-xs">
            <button
              onClick={() => setViewMode('both')}
              className={`px-3 py-1 rounded-md font-medium flex items-center gap-1.5 transition-all ${
                viewMode === 'both'
                  ? 'bg-[#2b579a] text-white shadow-xs'
                  : 'text-slate-300 hover:text-white hover:bg-slate-700'
              }`}
              title="Ver simulador de Word 2016 y panel lateral juntos"
            >
              <Monitor className="w-3.5 h-3.5" />
              <span className="hidden md:inline">Word Simulado + Panel</span>
            </button>

            <button
              onClick={() => setViewMode('taskpane_only')}
              className={`px-3 py-1 rounded-md font-medium flex items-center gap-1.5 transition-all ${
                viewMode === 'taskpane_only'
                  ? 'bg-[#2b579a] text-white shadow-xs'
                  : 'text-slate-300 hover:text-white hover:bg-slate-700'
              }`}
              title="Ver exclusivamente el panel lateral exactamente como se ve en Word"
            >
              <PanelRight className="w-3.5 h-3.5" />
              <span className="hidden md:inline">Solo Panel Lateral</span>
            </button>
          </div>

          {/* Sideloading & Manifest XML button */}
          <button
            onClick={() => setIsManifestModalOpen(true)}
            className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg font-semibold text-xs flex items-center gap-1.5 shadow-sm transition-all"
            title="Descargar archivo manifest.xml para Word 2016 de escritorio"
          >
            <FileCode className="w-3.5 h-3.5" />
            <span>Instalar en Word 2016 (XML)</span>
          </button>
        </div>
      </header>

      {/* Main Workspace */}
      <main className="flex-1 flex overflow-hidden">
        {/* Word 2016 Document Simulator (Only visible in browser when 'both' is active) */}
        {viewMode === 'both' && (
          <div className="flex-1 min-w-0 h-full overflow-hidden">
            <WordSimulator
              documentText={documentText}
              onDocumentTextChange={(val) => {
                setDocumentText(val);
                const scanned = scanDocumentTextIntoBlocks(val, titleKeyword);
                setBlocks((prev) => mergeBlocksPreservingRatings(scanned, prev));
              }}
              titleKeyword={titleKeyword}
              blocks={blocks}
              activeBlockIndex={activeBlockIndex}
              onBlockSelect={handleSelectBlock}
              fileName={filePath.split(/[\\/]/).pop() || 'Documento1.docx'}
              onInsertBlockBetween={() => {
                const lines = documentText.split('\n');
                const insertPosition = lines.findIndex((l) =>
                  l.toLowerCase().includes(`${titleKeyword.toLowerCase()} 2`)
                );
                const newBlockText = `\n${titleKeyword} 1.5 (Nuevo en medio): Javier Ortiz - Consultor de Ciberseguridad\nNuevo bloque insertado entre el bloque 1 y el bloque 2 para verificar que los bloques existentes conservan sus calificaciones intactas sin alteración.\n`;
                let newDoc = '';
                if (insertPosition !== -1) {
                  lines.splice(insertPosition, 0, newBlockText);
                  newDoc = lines.join('\n');
                } else {
                  newDoc = documentText + '\n' + newBlockText;
                }
                setDocumentText(newDoc);
                const scanned = scanDocumentTextIntoBlocks(newDoc, titleKeyword);
                setBlocks((prev) => mergeBlocksPreservingRatings(scanned, prev));
              }}
              onInsertBlockEnd={() => {
                const nextNum = blocks.length + 1;
                const appendText = `\n\n${titleKeyword} ${nextNum}: Sofía Castro - Administradora de Base de Datos\nEspecialista en replicación de datos, alta disponibilidad y optimización de consultas SQL complejas para informes de analítica en tiempo real.\n`;
                const newDoc = documentText + appendText;
                setDocumentText(newDoc);
                const scanned = scanDocumentTextIntoBlocks(newDoc, titleKeyword);
                setBlocks((prev) => mergeBlocksPreservingRatings(scanned, prev));
              }}
              onResetSampleDocument={() => {
                setDocumentText(INITIAL_SAMPLE_TEXT);
                setTitleKeyword('BLOQUE');
                const scanned = scanDocumentTextIntoBlocks(INITIAL_SAMPLE_TEXT, 'BLOQUE');
                setBlocks(
                  scanned.map((b) => ({
                    ...b,
                    ratings: {
                      datos: 'red',
                      experiencia: 'red',
                      tareas: 'red',
                    },
                  }))
                );
                setActiveBlockIndex(1);
              }}
            />
          </div>
        )}

        {/* Taskpane (The Web Add-in Lateral Panel) */}
        <div
          className={`${
            viewMode === 'taskpane_only'
              ? 'w-full max-w-md mx-auto my-0 h-full border-x border-slate-700 shadow-2xl'
              : 'w-[360px] lg:w-[380px] shrink-0 h-full'
          }`}
        >
          <Taskpane
            titleKeyword={titleKeyword}
            onTitleKeywordChange={setTitleKeyword}
            onSegmentDocument={handleSegmentDocument}
            blocks={blocks}
            activeBlockIndex={activeBlockIndex}
            onSelectBlock={handleSelectBlock}
            onUpdateRating={handleUpdateRating}
            filePath={filePath}
            onFilePathChange={setFilePath}
            isWordHost={isWordHost}
            onOpenManifestGuide={() => setIsManifestModalOpen(true)}
          />
        </div>
      </main>

      {/* Manifest & Installation Modal */}
      <ManifestModal
        isOpen={isManifestModalOpen}
        onClose={() => setIsManifestModalOpen(false)}
        appUrl={window.location.origin}
      />
    </div>
  );
}
