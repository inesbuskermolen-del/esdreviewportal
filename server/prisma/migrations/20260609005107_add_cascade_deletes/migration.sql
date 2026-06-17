-- DropForeignKey
ALTER TABLE "Credit" DROP CONSTRAINT "Credit_projectId_fkey";

-- DropForeignKey
ALTER TABLE "CreditComment" DROP CONSTRAINT "CreditComment_creditId_fkey";

-- DropForeignKey
ALTER TABLE "DrawingRequirement" DROP CONSTRAINT "DrawingRequirement_projectId_fkey";

-- DropForeignKey
ALTER TABLE "ESDExcellenceNote" DROP CONSTRAINT "ESDExcellenceNote_excellenceId_fkey";

-- DropForeignKey
ALTER TABLE "ESDExcellenceOpportunity" DROP CONSTRAINT "ESDExcellenceOpportunity_projectId_fkey";

-- DropForeignKey
ALTER TABLE "Reviewer" DROP CONSTRAINT "Reviewer_projectId_fkey";

-- AddForeignKey
ALTER TABLE "Credit" ADD CONSTRAINT "Credit_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "CreditComment" ADD CONSTRAINT "CreditComment_creditId_fkey" FOREIGN KEY ("creditId") REFERENCES "Credit"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Reviewer" ADD CONSTRAINT "Reviewer_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "DrawingRequirement" ADD CONSTRAINT "DrawingRequirement_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "ESDExcellenceOpportunity" ADD CONSTRAINT "ESDExcellenceOpportunity_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "ESDExcellenceNote" ADD CONSTRAINT "ESDExcellenceNote_excellenceId_fkey" FOREIGN KEY ("excellenceId") REFERENCES "ESDExcellenceOpportunity"("id") ON DELETE CASCADE ON UPDATE CASCADE;
