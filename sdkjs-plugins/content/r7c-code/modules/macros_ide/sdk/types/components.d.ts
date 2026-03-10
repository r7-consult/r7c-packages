/**
 * @fileoverview UI Components Type Definitions
 * @description Types for chat interface, file manager, and document viewer components
 * @author OnlyOffice UI SDK Team
 * @version 1.0.0
 */

import { BaseComponent, BaseComponentOptions, OnlyOfficeTheme } from './core';

// Chat Interface Types
export interface ChatInterfaceOptions extends BaseComponentOptions {
  apiKey?: string;
  endpoint?: string;
  model?: string;
  maxTokens?: number;
  temperature?: number;
  systemPrompt?: string;
  allowMarkdown?: boolean;
  autoScroll?: boolean;
  maxMessages?: number;
  enableTyping?: boolean;
  showTimestamps?: boolean;
  allowFileUpload?: boolean;
  allowImageUpload?: boolean;
  placeholder?: string;
  sendButtonText?: string;
  clearButtonText?: string;
  retryButtonText?: string;
  maxRetries?: number;
  requestTimeout?: number;
}

export interface ChatMessage {
  id: string;
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: number;
  metadata?: Record<string, any>;
  attachments?: ChatAttachment[];
  status?: 'sending' | 'sent' | 'error' | 'retry';
  error?: string;
}

export interface ChatAttachment {
  id: string;
  name: string;
  type: string;
  size: number;
  url?: string;
  preview?: string;
  metadata?: Record<string, any>;
}

export interface SendMessageOptions {
  role?: 'user' | 'system';
  metadata?: Record<string, any>;
  attachments?: ChatAttachment[];
  stream?: boolean;
  onToken?: (token: string) => void;
  onProgress?: (progress: number) => void;
  signal?: AbortSignal;
}

export interface ChatInterfaceEvents {
  'message:sent': { message: ChatMessage };
  'message:received': { message: ChatMessage };
  'message:error': { message: ChatMessage; error: Error };
  'typing:start': { isAssistant: boolean };
  'typing:stop': { isAssistant: boolean };
  'conversation:cleared': {};
  'attachment:added': { attachment: ChatAttachment };
  'attachment:removed': { attachmentId: string };
  'stream:start': { messageId: string };
  'stream:token': { messageId: string; token: string };
  'stream:end': { messageId: string };
  'stream:error': { messageId: string; error: Error };
}

export declare class ChatInterface extends BaseComponent {
  constructor(options?: ChatInterfaceOptions);
  
  sendMessage(content: string, options?: SendMessageOptions): Promise<ChatMessage>;
  retryMessage(messageId: string): Promise<ChatMessage>;
  editMessage(messageId: string, newContent: string): Promise<ChatMessage>;
  deleteMessage(messageId: string): Promise<void>;
  
  getMessages(): ChatMessage[];
  getMessage(id: string): ChatMessage | null;
  clearMessages(): void;
  
  setSystemPrompt(prompt: string): void;
  getSystemPrompt(): string;
  
  getConversationContext(): string;
  exportConversation(format: 'json' | 'txt' | 'html'): string;
  importConversation(data: string, format: 'json'): void;
  
  addAttachment(attachment: ChatAttachment): void;
  removeAttachment(attachmentId: string): void;
  getAttachments(): ChatAttachment[];
  
  startTyping(): void;
  stopTyping(): void;
  isTyping(): boolean;
  
  scrollToBottom(): void;
  scrollToMessage(messageId: string): void;
  
  enableAutoScroll(): void;
  disableAutoScroll(): void;
  
  setPlaceholder(text: string): void;
  getPlaceholder(): string;
  
  setMaxMessages(max: number): void;
  getMaxMessages(): number;
  
  abortCurrentRequest(): void;
  isRequestInProgress(): boolean;
  
  applyTheme(theme: OnlyOfficeTheme): void;
}

// File Manager Types
export interface FileManagerOptions extends BaseComponentOptions {
  allowedTypes?: string[];
  maxFileSize?: number;
  allowMultiple?: boolean;
  showPreview?: boolean;
  enableUpload?: boolean;
  uploadEndpoint?: string;
  downloadEndpoint?: string;
  deleteEndpoint?: string;
  enableDragDrop?: boolean;
  enableContextMenu?: boolean;
  viewMode?: 'list' | 'grid' | 'tiles';
  sortBy?: 'name' | 'size' | 'date' | 'type';
  sortOrder?: 'asc' | 'desc';
  enableSearch?: boolean;
  enableFilters?: boolean;
  maxFiles?: number;
  chunkSize?: number;
  enableResumableUpload?: boolean;
  enableVersioning?: boolean;
}

export interface FileItem {
  id: string;
  name: string;
  type: string;
  mimeType: string;
  size: number;
  lastModified: number;
  created?: number;
  url?: string;
  downloadUrl?: string;
  preview?: string;
  thumbnail?: string;
  metadata?: Record<string, any>;
  permissions?: FilePermissions;
  version?: number;
  versions?: FileVersion[];
  status?: 'uploading' | 'uploaded' | 'error' | 'processing';
  progress?: number;
  error?: string;
  checksum?: string;
  tags?: string[];
}

export interface FilePermissions {
  read: boolean;
  write: boolean;
  delete: boolean;
  share: boolean;
  download: boolean;
}

export interface FileVersion {
  version: number;
  size: number;
  lastModified: number;
  url: string;
  checksum: string;
  comment?: string;
}

export interface FileFilter {
  type?: string[];
  size?: { min?: number; max?: number };
  date?: { from?: Date; to?: Date };
  name?: string;
  tags?: string[];
}

export interface FileUploadOptions {
  onProgress?: (progress: number, file: FileItem) => void;
  onComplete?: (file: FileItem) => void;
  onError?: (error: Error, file: FileItem) => void;
  metadata?: Record<string, any>;
  tags?: string[];
  replace?: boolean;
  resumable?: boolean;
  chunkSize?: number;
}

export interface FileManagerEvents {
  'file:added': { file: FileItem };
  'file:removed': { fileId: string };
  'file:selected': { file: FileItem };
  'file:deselected': { file: FileItem };
  'file:uploaded': { file: FileItem };
  'file:download': { file: FileItem };
  'file:preview': { file: FileItem };
  'file:error': { file: FileItem; error: Error };
  'upload:start': { files: FileItem[] };
  'upload:progress': { file: FileItem; progress: number };
  'upload:complete': { file: FileItem };
  'upload:error': { file: FileItem; error: Error };
  'selection:changed': { selectedFiles: FileItem[] };
  'view:changed': { viewMode: string };
  'sort:changed': { sortBy: string; sortOrder: string };
  'filter:changed': { filter: FileFilter };
}

export declare class FileManager extends BaseComponent {
  constructor(options?: FileManagerOptions);
  
  getFiles(): FileItem[];
  getFile(id: string): FileItem | null;
  addFile(file: File | FileItem): Promise<FileItem>;
  addFiles(files: (File | FileItem)[]): Promise<FileItem[]>;
  removeFile(id: string): Promise<void>;
  removeFiles(ids: string[]): Promise<void>;
  updateFile(id: string, updates: Partial<FileItem>): Promise<FileItem>;
  
  selectFile(id: string): void;
  deselectFile(id: string): void;
  selectAll(): void;
  deselectAll(): void;
  getSelectedFiles(): FileItem[];
  isFileSelected(id: string): boolean;
  
  uploadFiles(files: File[], options?: FileUploadOptions): Promise<FileItem[]>;
  uploadFile(file: File, options?: FileUploadOptions): Promise<FileItem>;
  pauseUpload(fileId: string): void;
  resumeUpload(fileId: string): void;
  cancelUpload(fileId: string): void;
  retryUpload(fileId: string): Promise<FileItem>;
  
  downloadFile(id: string): Promise<void>;
  downloadFiles(ids: string[]): Promise<void>;
  downloadSelected(): Promise<void>;
  
  previewFile(id: string): void;
  closePreview(): void;
  isPreviewOpen(): boolean;
  
  searchFiles(query: string): FileItem[];
  filterFiles(filter: FileFilter): FileItem[];
  clearFilter(): void;
  getActiveFilter(): FileFilter | null;
  
  setViewMode(mode: 'list' | 'grid' | 'tiles'): void;
  getViewMode(): string;
  
  setSortBy(field: 'name' | 'size' | 'date' | 'type'): void;
  setSortOrder(order: 'asc' | 'desc'): void;
  getSortBy(): string;
  getSortOrder(): string;
  
  refresh(): Promise<void>;
  
  enableDragDrop(): void;
  disableDragDrop(): void;
  isDragDropEnabled(): boolean;
  
  showContextMenu(fileId: string, x: number, y: number): void;
  hideContextMenu(): void;
  
  getFileVersions(id: string): FileVersion[];
  restoreFileVersion(id: string, version: number): Promise<FileItem>;
  
  addTag(fileId: string, tag: string): Promise<void>;
  removeTag(fileId: string, tag: string): Promise<void>;
  getFileTags(fileId: string): string[];
  getAllTags(): string[];
  
  getFilePermissions(id: string): FilePermissions;
  setFilePermissions(id: string, permissions: Partial<FilePermissions>): Promise<void>;
  
  getStorageInfo(): Promise<{
    used: number;
    available: number;
    total: number;
    fileCount: number;
  }>;
  
  applyTheme(theme: OnlyOfficeTheme): void;
}

// Document Viewer Types
export interface DocumentViewerOptions extends BaseComponentOptions {
  documentUrl?: string;
  documentType?: 'pdf' | 'docx' | 'xlsx' | 'pptx' | 'txt' | 'html' | 'md';
  enableAnnotations?: boolean;
  readonly?: boolean;
  zoom?: number;
  fitToWidth?: boolean;
  fitToHeight?: boolean;
  enableSearch?: boolean;
  enablePrint?: boolean;
  enableDownload?: boolean;
  enableFullscreen?: boolean;
  showToolbar?: boolean;
  showPageNumbers?: boolean;
  showNavigationPanel?: boolean;
  maxZoom?: number;
  minZoom?: number;
  zoomStep?: number;
  autoFit?: 'none' | 'width' | 'height' | 'page';
  scrollMode?: 'vertical' | 'horizontal' | 'wrapped';
  spreadMode?: 'none' | 'odd' | 'even';
  renderTextLayer?: boolean;
  enableTextSelection?: boolean;
  enableLinks?: boolean;
  cacheSize?: number;
  preloadPages?: number;
}

export interface DocumentInfo {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  creator?: string;
  producer?: string;
  creationDate?: Date;
  modificationDate?: Date;
  pageCount: number;
  fileSize?: number;
  version?: string;
  format: string;
  encrypted?: boolean;
  permissions?: DocumentPermissions;
}

export interface DocumentPermissions {
  print: boolean;
  copy: boolean;
  modify: boolean;
  annotate: boolean;
  fillForms: boolean;
  extract: boolean;
  assemble: boolean;
  printHighQuality: boolean;
}

export interface DocumentPage {
  pageNumber: number;
  width: number;
  height: number;
  rotation: number;
  content?: string;
  annotations?: DocumentAnnotation[];
  links?: DocumentLink[];
}

export interface DocumentAnnotation {
  id: string;
  type: 'highlight' | 'note' | 'strikeout' | 'underline' | 'squiggly' | 'circle' | 'square';
  page: number;
  bounds: { x: number; y: number; width: number; height: number };
  content?: string;
  author?: string;
  created: Date;
  modified?: Date;
  color?: string;
  opacity?: number;
}

export interface DocumentLink {
  bounds: { x: number; y: number; width: number; height: number };
  url?: string;
  page?: number;
  action?: 'goto' | 'uri' | 'javascript';
}

export interface SearchResult {
  page: number;
  text: string;
  bounds: { x: number; y: number; width: number; height: number };
  context: string;
}

export interface DocumentViewerEvents {
  'document:loaded': { document: DocumentInfo };
  'document:error': { error: Error };
  'page:changed': { pageNumber: number; totalPages: number };
  'zoom:changed': { zoom: number };
  'search:found': { results: SearchResult[]; query: string };
  'search:cleared': {};
  'annotation:added': { annotation: DocumentAnnotation };
  'annotation:modified': { annotation: DocumentAnnotation };
  'annotation:removed': { annotationId: string };
  'text:selected': { text: string; page: number };
  'link:clicked': { link: DocumentLink };
  'fullscreen:enter': {};
  'fullscreen:exit': {};
  'print:start': {};
  'print:complete': {};
  'download:start': {};
  'download:complete': {};
}

export declare class DocumentViewer extends BaseComponent {
  constructor(options?: DocumentViewerOptions);
  
  loadDocument(url: string, type?: string): Promise<void>;
  unloadDocument(): void;
  reloadDocument(): Promise<void>;
  
  getDocument(): DocumentInfo | null;
  isDocumentLoaded(): boolean;
  
  getCurrentPage(): number;
  getTotalPages(): number;
  goToPage(page: number): void;
  nextPage(): void;
  previousPage(): void;
  firstPage(): void;
  lastPage(): void;
  
  getZoom(): number;
  setZoom(level: number): void;
  zoomIn(): void;
  zoomOut(): void;
  resetZoom(): void;
  fitToWidth(): void;
  fitToHeight(): void;
  fitToPage(): void;
  
  getRotation(): number;
  setRotation(degrees: number): void;
  rotateClockwise(): void;
  rotateCounterClockwise(): void;
  
  searchText(query: string, options?: {
    caseSensitive?: boolean;
    wholeWords?: boolean;
    backwards?: boolean;
    highlightAll?: boolean;
  }): SearchResult[];
  
  findNext(): SearchResult | null;
  findPrevious(): SearchResult | null;
  clearSearch(): void;
  
  selectText(page: number, bounds: { x: number; y: number; width: number; height: number }): string;
  getSelectedText(): string;
  clearSelection(): void;
  
  addAnnotation(annotation: Omit<DocumentAnnotation, 'id' | 'created'>): Promise<DocumentAnnotation>;
  updateAnnotation(id: string, updates: Partial<DocumentAnnotation>): Promise<DocumentAnnotation>;
  removeAnnotation(id: string): Promise<void>;
  getAnnotations(page?: number): DocumentAnnotation[];
  clearAnnotations(page?: number): void;
  
  exportAnnotations(format: 'json' | 'xfdf'): string;
  importAnnotations(data: string, format: 'json' | 'xfdf'): Promise<void>;
  
  print(options?: {
    pages?: number[];
    copies?: number;
    orientation?: 'portrait' | 'landscape';
    paperSize?: string;
    margins?: { top: number; right: number; bottom: number; left: number };
  }): Promise<void>;
  
  download(filename?: string): Promise<void>;
  
  enterFullscreen(): void;
  exitFullscreen(): void;
  toggleFullscreen(): void;
  isFullscreen(): boolean;
  
  showToolbar(): void;
  hideToolbar(): void;
  isToolbarVisible(): boolean;
  
  showNavigationPanel(): void;
  hideNavigationPanel(): void;
  isNavigationPanelVisible(): boolean;
  
  setScrollMode(mode: 'vertical' | 'horizontal' | 'wrapped'): void;
  getScrollMode(): string;
  
  setSpreadMode(mode: 'none' | 'odd' | 'even'): void;
  getSpreadMode(): string;
  
  getPageInfo(page: number): DocumentPage | null;
  getVisiblePages(): number[];
  
  scrollToPage(page: number): void;
  scrollToPosition(x: number, y: number): void;
  getScrollPosition(): { x: number; y: number };
  
  enableTextSelection(): void;
  disableTextSelection(): void;
  isTextSelectionEnabled(): boolean;
  
  enableAnnotations(): void;
  disableAnnotations(): void;
  areAnnotationsEnabled(): boolean;
  
  getPerformanceMetrics(): {
    loadTime: number;
    renderTime: number;
    memoryUsage: number;
    cacheHitRate: number;
  };
  
  applyTheme(theme: OnlyOfficeTheme): void;
}

// Component Factory Functions
export function createChatInterface(options?: ChatInterfaceOptions): Promise<ChatInterface>;
export function createFileManager(options?: FileManagerOptions): Promise<FileManager>;
export function createDocumentViewer(options?: DocumentViewerOptions): Promise<DocumentViewer>;

// Component Registration
export interface ComponentRegistry {
  register<T extends BaseComponent>(
    name: string, 
    factory: (options?: any) => Promise<T>
  ): void;
  
  unregister(name: string): void;
  create<T extends BaseComponent>(name: string, options?: any): Promise<T>;
  isRegistered(name: string): boolean;
  getRegisteredComponents(): string[];
}

export const componentRegistry: ComponentRegistry;

export default {
  ChatInterface,
  FileManager,
  DocumentViewer,
  createChatInterface,
  createFileManager,
  createDocumentViewer,
  componentRegistry
};