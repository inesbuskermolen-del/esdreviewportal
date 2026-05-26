-- CreateTable
CREATE TABLE "ESDExcellenceNote" (
    "id" TEXT NOT NULL,
    "excellenceId" TEXT NOT NULL,
    "reviewerEmail" TEXT NOT NULL,
    "notes" TEXT NOT NULL DEFAULT '',
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "ESDExcellenceNote_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "ESDExcellenceNote_excellenceId_reviewerEmail_key" ON "ESDExcellenceNote"("excellenceId", "reviewerEmail");

-- AddForeignKey
ALTER TABLE "ESDExcellenceNote" ADD CONSTRAINT "ESDExcellenceNote_excellenceId_fkey" FOREIGN KEY ("excellenceId") REFERENCES "ESDExcellenceOpportunity"("id") ON DELETE RESTRICT ON UPDATE CASCADE;
