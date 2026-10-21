/**
 * Service to interface with Microsoft Word 2016 via Office.js APIs (WordApi 1.1 / 1.2 / 1.3)
 * Provides automatic detection of Word environment and document settings persistence.
 */

import { VirtualBlock, DocumentConfiguration, BlockRatings } from '../types';

declare const Office: any;
declare const Word: any;

export interface WordEnvironmentStatus {
  isOfficeHost: boolean;
  hostName: string | null;
  platform: string | null;
  documentUrl: string | null;
}

export class WordService {
  private static instance: WordService;
  private isInitialized = false;
  private isWordHost = false;
  private hostInfo: any = null;
  private documentUrl: string = '';

  private constructor() {}

  public static getInstance(): WordService {
    if (!WordService.instance) {
      WordService.instance = new WordService();
    }
    return WordService.instance;
  }

  /**
   * Initializes Office.js runtime safely without blocking the UI
   */
  public async initialize(): Promise<WordEnvironmentStatus> {
    if (this.isInitialized) {
      return this.getStatus();
    }

    try {
      // Dynamically load Office.js if not already present
      await this.loadOfficeScript();

      if (typeof Office !== 'undefined' && Office.onReady) {
        // Race with a 800ms timeout so standalone preview never hangs
        const timeoutPromise = new Promise<{ host: null; platform: null }>((resolve) =>
          setTimeout(() => resolve({ host: null, platform: null }), 800)
        );

        const info: any = await Promise.race([
          Office.onReady(),
          timeoutPromise,
        ]);

        this.hostInfo = info;
        this.isWordHost = info && info.host === Office.HostType.Word;
        this.isInitialized = true;

        if (this.isWordHost && Office.context?.document) {
          this.documentUrl = Office.context.document.url || 'C:\\Documentos\\Documento_Word_2016.docx';
        }
      } else {
        this.isInitialized = true;
        this.isWordHost = false;
      }
    } catch (e) {
      console.warn('WordService running in standalone/simulator mode:', e);
      this.isInitialized = true;
      this.isWordHost = false;
    }

    return this.getStatus();
  }

  private loadOfficeScript(): Promise<boolean> {
    return new Promise((resolve) => {
      if (typeof (window as any).Office !== 'undefined') {
        resolve(true);
        return;
      }

      if (document.getElementById('office-js-script')) {
        resolve(false);
        return;
      }

      const script = document.createElement('script');
      script.id = 'office-js-script';
      script.src = 'https://appsforoffice.microsoft.com/lib/1/hosted/office.js';
      script.async = true;
      script.onload = () => resolve(true);
      script.onerror = () => resolve(false);

      // 1-second timeout guard
      setTimeout(() => resolve(false), 1000);
      document.head.appendChild(script);
    });
  }

  public getStatus(): WordEnvironmentStatus {
    return {
      isOfficeHost: this.isWordHost,
      hostName: this.hostInfo?.host ? String(this.hostInfo.host) : null,
      platform: this.hostInfo?.platform ? String(this.hostInfo.platform) : null,
      documentUrl: this.documentUrl || 'C:\\Documentos\\Evaluacion_Candidatos_Word2016.docx',
    };
  }

  /**
   * Saves the add-in settings directly into the Word document package (.docx).
   * In Word 2016, Office.context.document.settings is persisted in the document itself when the user saves!
   */
  public async saveToDocumentSettings(config: DocumentConfiguration): Promise<boolean> {
    if (!this.isWordHost || typeof Office === 'undefined' || !Office.context?.document?.settings) {
      return false;
    }

    return new Promise((resolve) => {
      try {
        const settings = Office.context.document.settings;
        settings.set('WordBlockRatingsConfig', JSON.stringify(config));
        settings.saveAsync((asyncResult: any) => {
          if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            resolve(true);
          } else {
            console.error('Error saving to Office document settings:', asyncResult.error);
            resolve(false);
          }
        });
      } catch (err) {
        console.error('Failed to set Word document settings:', err);
        resolve(false);
      }
    });
  }

  /**
   * Loads saved settings stored inside the Word document (.docx)
   */
  public async loadFromDocumentSettings(): Promise<DocumentConfiguration | null> {
    if (!this.isWordHost || typeof Office === 'undefined' || !Office.context?.document?.settings) {
      return null;
    }

    try {
      const settings = Office.context.document.settings;
      const raw = settings.get('WordBlockRatingsConfig');
      if (raw && typeof raw === 'string') {
        return JSON.parse(raw) as DocumentConfiguration;
      }
    } catch (err) {
      console.error('Failed to load Word document settings:', err);
    }
    return null;
  }

  /**
   * Scans a real Word document for the given title keyword using Word.run.
   * WordApi 1.1 search API is compatible with Word 2016.
   */
  public async scanRealWordDocument(titleKeyword: string): Promise<Omit<VirtualBlock, 'ratings'>[] | null> {
    if (!this.isWordHost || typeof Word === 'undefined') {
      return null;
    }

    try {
      return await Word.run(async (context: any) => {
        const body = context.document.body;
        const searchResults = body.search(titleKeyword, { matchCase: false, matchWholeWord: false });
        context.load(searchResults, 'text');
        
        await context.sync();

        if (searchResults.items.length === 0) {
          return [];
        }

        const blocks: Omit<VirtualBlock, 'ratings'>[] = [];
        
        // For each occurrence, get the paragraph text and preview
        for (let i = 0; i < searchResults.items.length; i++) {
          const item = searchResults.items[i];
          const paragraphs = item.paragraphs;
          context.load(paragraphs, 'text');
        }

        await context.sync();

        for (let i = 0; i < searchResults.items.length; i++) {
          const item = searchResults.items[i];
          const headingParagraph = item.paragraphs.items[0];
          const headingText = headingParagraph ? headingParagraph.text.trim() : item.text.trim();
          
          const blockId = `word_blk_${i + 1}_${headingText.slice(0, 15).replace(/[^a-zA-Z0-9]/g, '_')}`;
          
          blocks.push({
            id: blockId,
            index: i + 1,
            titleText: headingText || `${titleKeyword} #${i + 1}`,
            snippet: `Sección detectada en el documento de Word (Posición #${i + 1})`,
          });
        }

        return blocks;
      });
    } catch (err) {
      console.error('Error scanning real Word document:', err);
      return null;
    }
  }

  /**
   * Selects and scrolls to a specific block in the real Word document
   */
  public async navigateToBlockInWord(titleKeyword: string, blockIndex: number): Promise<boolean> {
    if (!this.isWordHost || typeof Word === 'undefined') {
      return false;
    }

    try {
      return await Word.run(async (context: any) => {
        const body = context.document.body;
        const searchResults = body.search(titleKeyword, { matchCase: false, matchWholeWord: false });
        context.load(searchResults, 'text');
        await context.sync();

        const target = searchResults.items[blockIndex - 1];
        if (target) {
          // Selects the block in Word, which automatically scrolls Word's viewport to it
          target.select();
          await context.sync();
          return true;
        }
        return false;
      });
    } catch (err) {
      console.error('Error navigating in Word document:', err);
      return false;
    }
  }

  /**
   * Listens to cursor/selection changes in Word to detect which block the user is currently viewing/editing.
   * Uses WordApi 1.1 range.compareLocationWith() to accurately locate the cursor between block headings.
   */
  public subscribeToSelectionChanges(
    titleKeyword: string,
    onActiveBlockDetected: (blockIndex: number) => void
  ): () => void {
    if (!this.isWordHost || typeof Office === 'undefined' || !Office.context?.document) {
      return () => {};
    }

    let isChecking = false;

    const handler = () => {
      if (isChecking || !titleKeyword.trim()) return;
      isChecking = true;

      Word.run(async (context: any) => {
        const selection = context.document.getSelection();
        const body = context.document.body;
        const searchResults = body.search(titleKeyword, { matchCase: false, matchWholeWord: false });
        
        context.load(selection);
        context.load(searchResults, 'text');
        await context.sync();

        if (searchResults.items.length === 0) {
          isChecking = false;
          return;
        }

        // Compare location of each block heading relative to the user's current selection
        // In WordApi 1.1: item.compareLocationWith(selection)
        const relations: any[] = [];
        for (let i = 0; i < searchResults.items.length; i++) {
          const rel = searchResults.items[i].compareLocationWith(selection);
          relations.push(rel);
        }

        await context.sync();

        let detectedIndex = 1;

        for (let i = 0; i < relations.length; i++) {
          const val = relations[i].value;
          // 'Before' means the block header is located before the cursor
          // 'Inside', 'Contains', 'Equal' means the cursor is directly on the header
          if (val === 'Before' || val === 'Contains' || val === 'Inside' || val === 'Equal') {
            detectedIndex = i + 1;
          } else if (val === 'After') {
            // Reached a block header that is after the cursor
            break;
          }
        }

        onActiveBlockDetected(detectedIndex);
      })
        .catch((e: any) => {
          console.debug('Selection compare note:', e);
        })
        .finally(() => {
          isChecking = false;
        });
    };

    try {
      Office.context.document.addHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        handler
      );

      // Trigger one initial check
      handler();

      return () => {
        try {
          Office.context.document.removeHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            { handler }
          );
        } catch (e) {
          // ignore
        }
      };
    } catch (e) {
      return () => {};
    }
  }
}

export const wordService = WordService.getInstance();
