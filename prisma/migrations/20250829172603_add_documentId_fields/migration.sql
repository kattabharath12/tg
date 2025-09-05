
-- AlterTable
ALTER TABLE "IncomeEntry" ADD COLUMN "documentId" TEXT;

-- AlterTable
ALTER TABLE "DeductionEntry" ADD COLUMN "documentId" TEXT;

-- AddForeignKey
ALTER TABLE "IncomeEntry" ADD CONSTRAINT "IncomeEntry_documentId_fkey" FOREIGN KEY ("documentId") REFERENCES "Document"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "DeductionEntry" ADD CONSTRAINT "DeductionEntry_documentId_fkey" FOREIGN KEY ("documentId") REFERENCES "Document"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- CreateIndex
CREATE INDEX "IncomeEntry_documentId_idx" ON "IncomeEntry"("documentId");

-- CreateIndex
CREATE INDEX "DeductionEntry_documentId_idx" ON "DeductionEntry"("documentId");
