'use client'

import { useState, useRef } from 'react'
import { Button } from "@/components/ui/button"
import { Progress } from "@/components/ui/progress"
import { Card, CardContent, CardFooter, CardHeader, CardTitle } from "@/components/ui/card"
import { Alert, AlertDescription } from "@/components/ui/alert"
import { Upload, FileText, X, Check } from 'lucide-react'

export default function Component() {
  const [files, setFiles] = useState<File[]>([])
  const [error, setError] = useState<string | null>(null)
  const [converting, setConverting] = useState(false)
  const [progress, setProgress] = useState(0)
  const [converted, setConverted] = useState(false)
  const fileInputRef = useRef<HTMLInputElement>(null)

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const selectedFiles = Array.from(event.target.files || [])
    processFiles(selectedFiles)
  }

  const handleDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault()
    const droppedFiles = Array.from(event.dataTransfer.files)
    processFiles(droppedFiles)
  }

  const processFiles = (newFiles: File[]) => {
    setError(null)
    const validFiles = newFiles.filter(file => file.name.endsWith('.pptx'))
    const invalidFiles = newFiles.filter(file => !file.name.endsWith('.pptx'))

    if (invalidFiles.length > 0) {
      setError(`Invalid file(s): ${invalidFiles.map(f => f.name).join(', ')}. Only .pptx files are allowed.`)
    }

    setFiles(prevFiles => {
      const updatedFiles = [...prevFiles]
      validFiles.forEach(file => {
        if (!updatedFiles.some(f => f.name === file.name)) {
          updatedFiles.push(file)
        }
      })
      return updatedFiles
    })
  }

  const removeFile = (index: number) => {
    setFiles(prevFiles => prevFiles.filter((_, i) => i !== index))
  }

  const handleConvert = () => {
    setConverting(true)
    setProgress(0)
    const interval = setInterval(() => {
      setProgress(oldProgress => {
        if (oldProgress >= 100) {
          clearInterval(interval)
          setConverting(false)
          setConverted(true)
          return 100
        }
        return oldProgress + 5
      })
    }, 200)
  }

  const resetConverter = () => {
    setFiles([])
    setConverted(false)
    setProgress(0)
  }

  return (
    <Card className="w-full max-w-md mx-auto">
      <CardHeader>
        <CardTitle className="text-2xl font-bold text-center">PPT to PDF Converter</CardTitle>
      </CardHeader>
      <CardContent>
        <div
          className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center cursor-pointer hover:border-gray-400 transition-colors"
          onClick={() => fileInputRef.current?.click()}
          onDragOver={(e) => e.preventDefault()}
          onDrop={handleDrop}
        >
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileChange}
            accept=".pptx"
            multiple
            className="hidden"
          />
          <Upload className="mx-auto h-12 w-12 text-gray-400" />
          <p className="mt-2 text-sm text-gray-600">Click to select or drag and drop PPTX files here</p>
        </div>
        {error && (
          <Alert variant="destructive" className="mt-4">
            <AlertDescription>{error}</AlertDescription>
          </Alert>
        )}
        {files.length > 0 && (
          <div className="mt-4 space-y-2">
            {files.map((file, index) => (
              <div key={index} className="flex items-center justify-between bg-gray-100 p-2 rounded">
                <span className="text-sm truncate">{file.name}</span>
                <Button variant="ghost" size="icon" onClick={() => removeFile(index)} disabled={converting}>
                  <X className="h-4 w-4" />
                </Button>
              </div>
            ))}
          </div>
        )}
        {converting && <Progress value={progress} className="mt-4" />}
      </CardContent>
      <CardFooter className="flex flex-col items-center">
        {!converted ? (
          <Button onClick={handleConvert} disabled={files.length === 0 || converting} className="w-full">
            {converting ? 'Converting...' : 'Convert to PDF'}
          </Button>
        ) : (
          <div className="text-center">
            <Check className="mx-auto h-8 w-8 text-green-500 mb-2" />
            <p className="text-sm text-gray-600 mb-2">Conversion complete!</p>
            <Button variant="outline" className="w-full mb-2">
              <FileText className="mr-2 h-4 w-4" />
              Download PDF
            </Button>
            <Button variant="ghost" onClick={resetConverter} className="w-full">
              Convert Another File
            </Button>
          </div>
        )}
      </CardFooter>
    </Card>
  )
}
