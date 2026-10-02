import { Component, inject, signal } from '@angular/core';
import { JsonPipe } from '@angular/common';
import { HttpClient } from '@angular/common/http';

interface ExtractedLayout {
  layoutId?: string;
  source?: string;
  structure?: {
    sectionCount?: number;
    topLevelTableCount?: number;
  };
}

@Component({
  selector: 'app-root',
  imports: [JsonPipe],
  templateUrl: './app.html',
  styleUrl: './app.css',
})
export class App {
  isGenerating = signal(false);
  generateMessage = signal('');

  generateDocument(): void {
  const reportLayout = this.layout();

  if (!reportLayout) {
    this.generateMessage.set('Extract the layout first.');
    return;
  }

  this.isGenerating.set(true);
  this.generateMessage.set('');

  this.http
    .post('/api/layout/generate', reportLayout, { responseType: 'blob' })
    .subscribe({
      next: (docxBlob) => {
        const downloadUrl = window.URL.createObjectURL(docxBlob);
        const link = document.createElement('a');

        link.href = downloadUrl;
        link.download = 'generated-layout.docx';
        link.click();

        window.URL.revokeObjectURL(downloadUrl);
        this.isGenerating.set(false);
        this.generateMessage.set('DOCX generated. The download has started.');
      },
      error: (error) => {
        console.error('[Frontend] DOCX generation failed:', error);
        this.isGenerating.set(false);
        this.generateMessage.set(
          'Could not generate the DOCX. Check the backend terminal for details.',
        );
      },
    });
}
  private readonly http = inject(HttpClient);

  selectedFile = signal<File | null>(null);
  layout = signal<ExtractedLayout | null>(null);
  isExtracting = signal(false);
  errorMessage = signal('');

  onFileSelected(event: Event): void {
    const input = event.target as HTMLInputElement;
    this.selectedFile.set(input.files?.[0] ?? null);
    this.layout.set(null);
    this.errorMessage.set('');
  }

  extractLayout(): void {
    const file = this.selectedFile();

    if (!file) {
      this.errorMessage.set('Please choose a DOCX file first.');
      return;
    }

    const formData = new FormData();
    formData.append('file', file);

    this.isExtracting.set(true);
    this.errorMessage.set('');

    this.http.post<ExtractedLayout>('/api/layout/extract', formData).subscribe({
      next: (layout) => {
        this.layout.set(layout);
        this.isExtracting.set(false);
      },
      error: (error) => {
        this.errorMessage.set(
          error.error?.error ?? 'Could not extract the document layout.',
        );
        this.isExtracting.set(false);
      },
    });
  }
}