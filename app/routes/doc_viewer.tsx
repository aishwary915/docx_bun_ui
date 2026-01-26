import { useState, useCallback, lazy, Suspense } from "react";
import { useDropzone } from "react-dropzone";
import type { ClientLoaderFunctionArgs } from "react-router";
import {
  Card,
  CardContent,
  CardHeader,
  CardTitle,
} from "../components/ui/card";
import { Button } from "../components/ui/button";
import {
  Upload,
  FileText,
  Image as ImageIcon,
  FileSpreadsheet,
  File,
  X,
  Download,
  Eye,
  FolderOpen,
  CheckCircle2,
  CloudUpload,
  Edit3,
} from "lucide-react";

// Lazy load Univer components to prevent SSR issues
const UniverDocEditor = lazy(() => import("~/components/ui/UniverDocEditor").then(m => ({ default: m.UniverDocEditor })));
const UniverXlsxEditor = lazy(() => import("~/components/ui/UniverXlsxEditor").then(m => ({ default: m.UniverXlsxEditor })));

// This ensures the route is only rendered on the client
export async function clientLoader({}: ClientLoaderFunctionArgs) {
  return null;
}

interface UploadedFile {
  file: File;
  preview?: string;
  type: string;
}

function DocViewer() {
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [editingFile, setEditingFile] = useState<File | null>(null);

  const onDrop = useCallback((acceptedFiles: File[]) => {
    const newFiles = acceptedFiles.map((file): UploadedFile => {
      const fileType = file.type.split("/")[0];
      return {
        file,
        preview: fileType === "image" ? URL.createObjectURL(file) : undefined,
        type: fileType || "unknown",
      };
    });
    setUploadedFiles((prev) => [...prev, ...newFiles]);
  }, []);

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      "image/*": [".png", ".jpg", ".jpeg", ".gif", ".webp"],
      "application/pdf": [".pdf"],
      "text/csv": [".csv"],
      "application/vnd.ms-excel": [".xls"],
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [
        ".xlsx",
      ],
      "application/msword": [".doc"],
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        [".docx"],
      "text/plain": [".txt"],
    },
  });

  const removeFile = (index: number) => {
    setUploadedFiles((prev) => {
      const newFiles = [...prev];
      if (newFiles[index]?.preview) {
        URL.revokeObjectURL(newFiles[index].preview!);
      }
      newFiles.splice(index, 1);
      return newFiles;
    });
  };

  const getFileIcon = (file: File) => {
    const extension = file.name.split(".").pop()?.toLowerCase();

    if (file.type.startsWith("image/")) {
      return <ImageIcon className="h-10 w-10 text-blue-500" />;
    }
    if (file.type === "application/pdf") {
      return <FileText className="h-10 w-10 text-red-500" />;
    }
    if (extension === "csv" || extension === "xlsx" || extension === "xls") {
      return <FileSpreadsheet className="h-10 w-10 text-emerald-500" />;
    }
    return <File className="h-10 w-10 text-violet-500" />;
  };

  const formatFileSize = (bytes: number) => {
    if (bytes === 0) return "0 Bytes";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + " " + sizes[i];
  };

  const downloadFile = (file: File) => {
    const url = URL.createObjectURL(file);
    const a = document.createElement("a");
    a.href = url;
    a.download = file.name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const editDocFile = (file: File) => {
    setEditingFile(file);
  };

  const closeEditor = () => {
    setEditingFile(null);
  };

  // Check file type for appropriate editor/icon
  const isXlsxFile = (file: File) => {
    return file.name.endsWith(".xlsx") || file.name.endsWith(".xls") || 
           file.type === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
           file.type === "application/vnd.ms-excel";
  };

  const isDocxFile = (file: File) => {
    return file.name.endsWith(".docx") || file.name.endsWith(".doc") ||
           file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
           file.type === "application/msword";
  };

  // If editing a file, show the appropriate Univer editor
  if (editingFile) {
    const fileUrl = URL.createObjectURL(editingFile);
    const isXlsx = isXlsxFile(editingFile);
    const isDocx = isDocxFile(editingFile);

    return (
      <div className="fixed inset-0 z-50 flex flex-col bg-background">
        {/* Header with file name and close button */}
        <div className="flex items-center justify-between border-b bg-background p-4 shadow-sm">
          <div className="flex items-center gap-3">
            {isXlsx ? (
              <FileSpreadsheet className="h-6 w-6 text-emerald-500" />
            ) : (
              <FileText className="h-6 w-6 text-primary" />
            )}
            <div>
              <h2 className="font-semibold text-lg">{editingFile.name}</h2>
              <p className="text-muted-foreground text-sm">
                Editing with Univer {isXlsx ? "Sheets" : "Docs"}
              </p>
            </div>
          </div>
          <Button variant="outline" onClick={closeEditor} className="gap-2">
            <X className="h-4 w-4" />
            Close Editor
          </Button>
        </div>

        {/* Univer Editor */}
        <div className="flex-1 overflow-hidden">
          <Suspense fallback={<div className="flex h-full items-center justify-center"><p>Loading editor...</p></div>}>
            {isXlsx ? (
              <UniverXlsxEditor url={fileUrl} filename={editingFile.name} />
            ) : isDocx ? (
              <UniverDocEditor initialFile={editingFile} />
            ) : (
              <div className="flex h-full items-center justify-center">
                <p className="text-muted-foreground">Unsupported file type for editing</p>
              </div>
            )}
          </Suspense>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <div className="bg-white border-b">
        <div className="mx-auto max-w-7xl px-6 py-8">
          <div className="text-center">
            <h1 className="text-4xl font-bold text-gray-900 mb-2">
              Universal Document Viewer
            </h1>
            <p className="text-lg text-gray-600">
              View and download PDFs, Excel files, DOCX documents, images, and
              more
            </p>
          </div>
        </div>
      </div>

      {/* Main Content */}
      <div className="mx-auto max-w-7xl px-6 py-12">
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
          {/* Main Upload Area - Left Side */}
          <div className="lg:col-span-2">
            <Card className="h-full">
              <CardContent className="p-8">
                <div
                  {...getRootProps()}
                  className={`relative cursor-pointer overflow-hidden rounded-xl border-2 border-dashed p-16 transition-all duration-300 ${
                    isDragActive
                      ? "border-blue-400 bg-blue-50"
                      : "border-gray-300 hover:border-gray-400 hover:bg-gray-50"
                  }`}
                >
                  <input {...getInputProps()} />

                  <div className="text-center">
                    {/* Upload Icon */}
                    <div className="mx-auto mb-6 flex h-16 w-16 items-center justify-center rounded-full bg-blue-100">
                      {isDragActive ? (
                        <CloudUpload className="h-8 w-8 text-blue-600" />
                      ) : (
                        <Upload className="h-8 w-8 text-blue-600" />
                      )}
                    </div>

                    {/* Upload Text */}
                    {isDragActive ? (
                      <p className="text-xl font-medium text-blue-600 mb-2">
                        Drop your files here
                      </p>
                    ) : (
                      <>
                        <p className="text-xl font-medium text-gray-900 mb-2">
                          No file selected
                        </p>
                        <p className="text-gray-600 mb-4">
                          Drag & drop a document here or click to upload.
                        </p>
                        <p className="text-gray-500 text-sm mb-6">
                          Supported formats include PDF, Excel, DOCX, CSV,
                          <br />
                          and images.
                        </p>
                      </>
                    )}

                    {/* Choose File Button */}
                    {!isDragActive && (
                      <Button 
                        variant="default" 
                        size="lg"
                        className="bg-gray-800 hover:bg-gray-900 text-white px-8 pointer-events-none"
                      >
                        <Upload className="mr-2 h-5 w-5" />
                        Choose File
                      </Button>
                    )}
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>

          {/* Sidebar - Right Side */}
          <div className="lg:col-span-1">
            <div className="space-y-8">
              {/* Get Started Section */}
              <Card>
                <CardHeader>
                  <CardTitle className="text-xl font-bold">
                    Get Started
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <Button
                    variant="default"
                    size="lg"
                    className="w-full bg-gray-800 hover:bg-gray-900 text-white"
                    onClick={() =>
                      (document.querySelector('input[type="file"]') as HTMLInputElement)?.click()
                    }
                  >
                    <Upload className="mr-2 h-5 w-5" />
                    Upload File
                  </Button>
                </CardContent>
              </Card>

              {/* Supported Formats */}
              <Card>
                <CardHeader>
                  <CardTitle className="text-lg font-semibold">
                    Supported Formats
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <div className="space-y-3">
                    <div className="flex items-start gap-3">
                      <FileText className="h-5 w-5 text-red-500 mt-0.5" />
                      <div>
                        <p className="font-medium text-sm">PDF .pdf</p>
                        <p className="text-xs text-gray-500">
                          View, download, annotate & highlight
                        </p>
                      </div>
                    </div>

                    <div className="flex items-start gap-3">
                      <FileSpreadsheet className="h-5 w-5 text-green-500 mt-0.5" />
                      <div>
                        <p className="font-medium text-sm">Excel .xlsx</p>
                        <p className="text-xs text-gray-500">
                          View, download & advanced features
                        </p>
                      </div>
                    </div>

                    <div className="flex items-start gap-3">
                      <File className="h-5 w-5 text-orange-500 mt-0.5" />
                      <div>
                        <p className="font-medium text-sm">CSV .csv</p>
                        <p className="text-xs text-gray-500">View & download</p>
                      </div>
                    </div>

                    <div className="flex items-start gap-3">
                      <FileText className="h-5 w-5 text-blue-500 mt-0.5" />
                      <div>
                        <p className="font-medium text-sm">Word .docx</p>
                        <p className="text-xs text-gray-500">
                          View, download & scratchpad
                        </p>
                      </div>
                    </div>

                    <div className="flex items-start gap-3">
                      <ImageIcon className="h-5 w-5 text-purple-500 mt-0.5" />
                      <div>
                        <p className="font-medium text-sm">
                          Images .jpg, .png, .gif, .svg, .webp
                        </p>
                        <p className="text-xs text-gray-500">View & download</p>
                      </div>
                    </div>
                  </div>
                </CardContent>
              </Card>

              {/* Tip */}
              <div className="rounded-lg bg-yellow-50 border border-yellow-200 p-4">
                <div className="flex items-start gap-2">
                  <div className="text-yellow-600">💡</div>
                  <div>
                    <p className="text-sm font-medium text-yellow-800">Tip:</p>
                    <p className="text-xs text-yellow-700 mt-1">
                      Drag and drop files anywhere on the viewer or click the
                      upload button to get started. Supports PDF, Excel, DOCX,
                      CSV, and image files.
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Uploaded Files Section */}
      {uploadedFiles.length > 0 && (
        <div className="mx-auto max-w-7xl px-6 py-8">
          <div className="mb-6 flex items-center justify-between">
            <h2 className="text-2xl font-bold">Uploaded Files</h2>
            <Button
              variant="outline"
              onClick={() => {
                uploadedFiles.forEach((item) => {
                  if (item.preview) URL.revokeObjectURL(item.preview);
                });
                setUploadedFiles([]);
              }}
            >
              Clear All
            </Button>
          </div>

          <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-3">
            {uploadedFiles.map((item, index) => (
              <Card key={index} className="group relative overflow-hidden">
                {/* Remove Button */}
                <Button
                  variant="destructive"
                  size="icon"
                  className="absolute right-2 top-2 z-10 h-8 w-8 opacity-0 group-hover:opacity-100 transition-opacity"
                  onClick={() => removeFile(index)}
                >
                  <X className="h-4 w-4" />
                </Button>

                {/* File Preview */}
                <div className="relative h-40 overflow-hidden bg-gray-100">
                  {item.preview ? (
                    <img
                      src={item.preview}
                      alt={item.file.name}
                      className="h-full w-full object-cover"
                    />
                  ) : (
                    <div className="flex h-full items-center justify-center">
                      {getFileIcon(item.file)}
                    </div>
                  )}
                </div>

                <CardContent className="p-4">
                  <h3 className="font-medium truncate mb-2">
                    {item.file.name}
                  </h3>
                  <p className="text-sm text-gray-500 mb-3">
                    {formatFileSize(item.file.size)}
                  </p>

                  <div className="grid grid-cols-2 gap-2">
                    {isDocxFile(item.file) || isXlsxFile(item.file) ? (
                      <>
                        <Button
                          size="sm"
                          onClick={() => editDocFile(item.file)}
                          className="text-xs"
                        >
                          <Edit3 className="mr-1 h-3 w-3" />
                          Edit
                        </Button>
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => downloadFile(item.file)}
                          className="text-xs"
                        >
                          <Download className="mr-1 h-3 w-3" />
                          Download
                        </Button>
                      </>
                    ) : (
                      <>
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => {
                            if (item.preview) {
                              window.open(item.preview, "_blank");
                            } else {
                              const url = URL.createObjectURL(item.file);
                              window.open(url, "_blank");
                              URL.revokeObjectURL(url);
                            }
                          }}
                          className="text-xs"
                        >
                          <Eye className="mr-1 h-3 w-3" />
                          View
                        </Button>
                        <Button
                          size="sm"
                          onClick={() => downloadFile(item.file)}
                          className="text-xs"
                        >
                          <Download className="mr-1 h-3 w-3" />
                          Download
                        </Button>
                      </>
                    )}
                  </div>
                </CardContent>
              </Card>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

export default DocViewer;
