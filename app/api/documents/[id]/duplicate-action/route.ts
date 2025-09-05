
import { NextResponse } from "next/server"
import { getServerSession } from "next-auth"
import { authOptions } from "@/lib/auth"
import { prisma } from "@/lib/db"

export const dynamic = "force-dynamic"

export async function POST(
  request: Request,
  { params }: { params: { id: string } }
) {
  console.log("🔍 [DUPLICATE_ACTION] Handling duplicate action for document:", params.id)
  
  try {
    const session = await getServerSession(authOptions)
    
    if (!session?.user?.email) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 })
    }

    const user = await prisma.user.findUnique({
      where: { email: session.user.email }
    })

    if (!user) {
      return NextResponse.json({ error: "User not found" }, { status: 404 })
    }

    // Parse the request body
    const { action, replacementDocumentId } = await request.json()
    console.log("🔍 [DUPLICATE_ACTION] Action:", action, "Replace document:", replacementDocumentId)

    // Find the document
    const document = await prisma.document.findFirst({
      where: { 
        id: params.id,
        taxReturn: {
          userId: user.id
        }
      }
    })

    if (!document) {
      return NextResponse.json({ error: "Document not found" }, { status: 404 })
    }

    switch (action) {
      case 'proceed':
        // Mark the document as verified and allow it to be imported
        console.log("✅ [DUPLICATE_ACTION] Proceeding with duplicate document")
        await prisma.document.update({
          where: { id: params.id },
          data: { 
            isVerified: true,
            verificationNotes: 'User approved duplicate document import'
          }
        })
        return NextResponse.json({ 
          success: true, 
          message: "Document approved for import despite duplicate warning" 
        })

      case 'cancel':
        // Mark the document as failed and don't import
        console.log("❌ [DUPLICATE_ACTION] Canceling duplicate document import")
        await prisma.document.update({
          where: { id: params.id },
          data: { 
            processingStatus: 'FAILED',
            verificationNotes: 'User cancelled due to duplicate warning'
          }
        })
        return NextResponse.json({ 
          success: true, 
          message: "Document import cancelled due to duplicate" 
        })

      case 'replace':
        if (!replacementDocumentId) {
          return NextResponse.json({ error: "Replacement document ID required" }, { status: 400 })
        }
        
        console.log("🔄 [DUPLICATE_ACTION] Replacing existing document:", replacementDocumentId)
        
        // Mark the old document as replaced
        await prisma.document.update({
          where: { id: replacementDocumentId },
          data: { 
            processingStatus: 'FAILED',
            verificationNotes: 'Replaced by newer document upload'
          }
        })

        // Mark the new document as verified
        await prisma.document.update({
          where: { id: params.id },
          data: { 
            isVerified: true,
            verificationNotes: 'Approved as replacement for duplicate document'
          }
        })

        return NextResponse.json({ 
          success: true, 
          message: "Document replaced successfully" 
        })

      default:
        return NextResponse.json({ error: "Invalid action" }, { status: 400 })
    }

  } catch (error) {
    console.error("💥 [DUPLICATE_ACTION] Error handling duplicate action:", error)
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    )
  }
}
