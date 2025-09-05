

import { NextResponse } from "next/server"
import { getServerSession } from "next-auth"
import { prisma } from "@/lib/db"

export const dynamic = "force-dynamic"

export async function POST(
  request: Request,
  { params }: { params: { id: string } }
) {
  try {
    const session = await getServerSession()
    
    if (!session?.user?.email) {
      return NextResponse.json({ error: "Unauthorized" }, { status: 401 })
    }

    const user = await prisma.user.findUnique({
      where: { email: session.user.email }
    })

    if (!user) {
      return NextResponse.json({ error: "User not found" }, { status: 404 })
    }

    const taxReturn = await prisma.taxReturn.findFirst({
      where: { 
        id: params.id,
        userId: user.id 
      }
    })

    if (!taxReturn) {
      return NextResponse.json({ error: "Tax return not found" }, { status: 404 })
    }

    const data = await request.json()

    // Validate documentId if provided
    if (data.documentId) {
      const document = await prisma.document.findFirst({
        where: {
          id: data.documentId,
          taxReturnId: params.id
        }
      })

      if (!document) {
        return NextResponse.json({ error: "Document not found" }, { status: 404 })
      }
    }

    const incomeEntry = await prisma.incomeEntry.create({
      data: {
        taxReturnId: params.id,
        incomeType: data.incomeType,
        amount: data.amount,
        description: data.description,
        employerName: data.employerName,
        employerEIN: data.employerEIN,
        payerName: data.payerName,
        payerTIN: data.payerTIN,
        federalTaxWithheld: data.federalTaxWithheld || 0,
        documentId: data.documentId || null, // Link to source document if provided
      }
    })

    return NextResponse.json(incomeEntry)
  } catch (error) {
    console.error("Error creating income entry:", error)
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    )
  }
}
