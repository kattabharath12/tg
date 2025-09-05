
import { NextResponse } from "next/server"
import { getServerSession } from "next-auth"
import { authOptions } from "@/lib/auth"
import { prisma } from "@/lib/db"
import { writeFile, mkdir } from "fs/promises"
import { join } from "path"
import { v4 as uuidv4 } from "uuid"

export const dynamic = "force-dynamic"

export async function POST(request: Request) {
  console.log("🔍 [UPLOAD] Starting document upload process...")
  
  try {
    console.log("🔍 [UPLOAD] Step 1: Getting server session...")
    const session = await getServerSession(authOptions)
    
    if (!session?.user?.email) {
      console.log("❌ [UPLOAD] No session or email found")
      return NextResponse.json({ 
        error: "Authentication required. Please log in to upload documents." 
      }, { status: 401 })
    }
    console.log("✅ [UPLOAD] Session found for email:", session.user.email)

    console.log("🔍 [UPLOAD] Step 2: Finding user in database...")
    const user = await prisma.user.findUnique({
      where: { email: session.user.email }
    })

    if (!user) {
      console.log("❌ [UPLOAD] User not found in database")
      return NextResponse.json({ error: "User not found" }, { status: 404 })
    }
    console.log("✅ [UPLOAD] User found:", user.id)

    console.log("🔍 [UPLOAD] Step 3: Parsing form data...")
    const formData = await request.formData()
    const file = formData.get('file') as File
    const taxReturnId = formData.get('taxReturnId') as string

    console.log("🔍 [UPLOAD] Form data parsed:", {
      fileName: file?.name,
      fileType: file?.type,
      fileSize: file?.size,
      taxReturnId
    })

    if (!file || !taxReturnId) {
      console.log("❌ [UPLOAD] Missing file or tax return ID")
      return NextResponse.json({ error: "Missing file or tax return ID" }, { status: 400 })
    }

    console.log("🔍 [UPLOAD] Step 4: Verifying tax return ownership...")
    const taxReturn = await prisma.taxReturn.findFirst({
      where: { 
        id: taxReturnId,
        userId: user.id 
      }
    })

    if (!taxReturn) {
      console.log("❌ [UPLOAD] Tax return not found or not owned by user")
      return NextResponse.json({ error: "Tax return not found" }, { status: 404 })
    }
    console.log("✅ [UPLOAD] Tax return verified:", taxReturn.id)

    console.log("🔍 [UPLOAD] Step 5: Validating file type and size...")
    const allowedTypes = ['application/pdf', 'image/png', 'image/jpeg', 'image/tiff']
    if (!allowedTypes.includes(file.type)) {
      console.log("❌ [UPLOAD] Invalid file type:", file.type)
      return NextResponse.json({ 
        error: "Invalid file type. Supported types: PDF, PNG, JPEG, TIFF" 
      }, { status: 400 })
    }

    if (file.size > 10 * 1024 * 1024) { // 10MB limit
      console.log("❌ [UPLOAD] File size exceeds limit:", file.size)
      return NextResponse.json({ 
        error: "File size exceeds 10MB limit" 
      }, { status: 400 })
    }
    console.log("✅ [UPLOAD] File validation passed")

    console.log("🔍 [UPLOAD] Step 6: Creating upload directory...")
    const uploadDir = join(process.cwd(), 'uploads', 'documents')
    await mkdir(uploadDir, { recursive: true })
    console.log("✅ [UPLOAD] Upload directory created/verified:", uploadDir)

    console.log("🔍 [UPLOAD] Step 7: Generating unique filename...")
    const fileExtension = file.name.split('.').pop()
    const uniqueFileName = `${uuidv4()}.${fileExtension}`
    const filePath = join(uploadDir, uniqueFileName)
    console.log("✅ [UPLOAD] File path generated:", filePath)

    console.log("🔍 [UPLOAD] Step 8: Saving file to filesystem...")
    const bytes = await file.arrayBuffer()
    const buffer = Buffer.from(bytes)
    await writeFile(filePath, buffer)
    console.log("✅ [UPLOAD] File saved to filesystem")

    console.log("🔍 [UPLOAD] Step 9: Determining document type...")
    const documentType = determineDocumentType(file.name)
    console.log("✅ [UPLOAD] Document type determined:", documentType)

    console.log("🔍 [UPLOAD] Step 10: Creating document record in database...")
    const document = await prisma.document.create({
      data: {
        taxReturnId,
        fileName: file.name,
        fileType: file.type,
        fileSize: file.size,
        filePath: filePath,
        documentType,
        processingStatus: 'PENDING'
      }
    })
    console.log("✅ [UPLOAD] Document record created:", document.id)

    console.log("🎉 [UPLOAD] Upload process completed successfully!")
    return NextResponse.json(document)
  } catch (error) {
    console.error("💥 [UPLOAD] Document upload error:", error)
    console.error("💥 [UPLOAD] Error stack:", error instanceof Error ? error.stack : 'No stack trace')
    
    // Return more specific error information in development
    if (process.env.NODE_ENV === 'development') {
      return NextResponse.json({
        error: "Internal server error",
        details: error instanceof Error ? error.message : 'Unknown error',
        stack: error instanceof Error ? error.stack : 'No stack trace'
      }, { status: 500 })
    }
    
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    )
  }
}

function determineDocumentType(fileName: string): any {
  const lowerName = fileName.toLowerCase()
  
  if (lowerName.includes('w-2') || lowerName.includes('w2')) {
    return 'W2'
  }
  if (lowerName.includes('1099-int') || lowerName.includes('1099int')) {
    return 'FORM_1099_INT'
  }
  if (lowerName.includes('1099-div') || lowerName.includes('1099div')) {
    return 'FORM_1099_DIV'
  }
  if (lowerName.includes('1099-misc') || lowerName.includes('1099misc')) {
    return 'FORM_1099_MISC'
  }
  if (lowerName.includes('1099-nec') || lowerName.includes('1099nec')) {
    return 'FORM_1099_NEC'
  }
  if (lowerName.includes('1099-r') || lowerName.includes('1099r')) {
    return 'FORM_1099_R'
  }
  if (lowerName.includes('1099-g') || lowerName.includes('1099g')) {
    return 'FORM_1099_G'
  }
  if (lowerName.includes('1099')) {
    // For generic 1099 files, default to MISC type
    return 'FORM_1099_MISC'
  }
  
  return 'UNKNOWN'
}
