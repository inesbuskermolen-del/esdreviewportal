-- CreateTable
CREATE TABLE "Project" (
    "id" TEXT NOT NULL,
    "name" TEXT NOT NULL,
    "address" TEXT,
    "bessScore" DOUBLE PRECISION,
    "date" TIMESTAMP(3),
    "revision" TEXT,
    "projectId" TEXT,
    "reviewLinkToken" TEXT NOT NULL,
    "generationStatus" TEXT NOT NULL DEFAULT 'idle',
    "notifyEmail" TEXT,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "parentProjectId" TEXT,

    CONSTRAINT "Project_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Credit" (
    "id" TEXT NOT NULL,
    "projectId" TEXT NOT NULL,
    "category" TEXT NOT NULL,
    "categoryOrder" INTEGER NOT NULL,
    "creditId" TEXT NOT NULL,
    "creditName" TEXT NOT NULL,
    "creditRequirement" TEXT,
    "mandatory" BOOLEAN NOT NULL DEFAULT false,
    "responsibleParty" TEXT,
    "creditStatus" TEXT NOT NULL,
    "creditScore" DOUBLE PRECISION,
    "creditWeight" DOUBLE PRECISION,
    "commentsGIW" TEXT,
    "scopedOutReason" TEXT,
    "rawDataPoints" TEXT,
    "lastEditedBy" TEXT,
    "lastEditedAt" TIMESTAMP(3),
    "deletedByGIW" BOOLEAN NOT NULL DEFAULT false,
    "hiddenFromPortal" BOOLEAN NOT NULL DEFAULT false,

    CONSTRAINT "Credit_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "CreditComment" (
    "id" TEXT NOT NULL,
    "creditId" TEXT NOT NULL,
    "projectId" TEXT NOT NULL,
    "reviewerEmail" TEXT NOT NULL,
    "reviewerDiscipline" TEXT NOT NULL,
    "commentText" TEXT NOT NULL,
    "submittedAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "CreditComment_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "Reviewer" (
    "id" TEXT NOT NULL,
    "projectId" TEXT NOT NULL,
    "email" TEXT NOT NULL,
    "discipline" TEXT NOT NULL,
    "hasSubmitted" BOOLEAN NOT NULL DEFAULT false,
    "submittedAt" TIMESTAMP(3),
    "inviteToken" TEXT,

    CONSTRAINT "Reviewer_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "DrawingRequirement" (
    "id" TEXT NOT NULL,
    "projectId" TEXT NOT NULL,
    "creditReference" TEXT NOT NULL,
    "drawingType" TEXT NOT NULL,
    "requirement" TEXT NOT NULL,
    "discipline" TEXT,
    "status" TEXT NOT NULL DEFAULT 'NotStarted',
    "notes" TEXT,

    CONSTRAINT "DrawingRequirement_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "ESDExcellenceOpportunity" (
    "id" TEXT NOT NULL,
    "projectId" TEXT NOT NULL,
    "creditId" TEXT,
    "creditReference" TEXT NOT NULL,
    "creditName" TEXT NOT NULL,
    "currentScore" DOUBLE PRECISION,
    "improvementDescription" TEXT,
    "flag" TEXT NOT NULL DEFAULT 'Unflagged',
    "flaggedBy" TEXT,
    "flaggedAt" TIMESTAMP(3),
    "reviewerNotes" TEXT,
    "bessPoints" TEXT,
    "deletedByGIW" BOOLEAN NOT NULL DEFAULT false,

    CONSTRAINT "ESDExcellenceOpportunity_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE UNIQUE INDEX "Project_reviewLinkToken_key" ON "Project"("reviewLinkToken");

-- CreateIndex
CREATE UNIQUE INDEX "Reviewer_inviteToken_key" ON "Reviewer"("inviteToken");

-- AddForeignKey
ALTER TABLE "Project" ADD CONSTRAINT "Project_parentProjectId_fkey" FOREIGN KEY ("parentProjectId") REFERENCES "Project"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Credit" ADD CONSTRAINT "Credit_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "CreditComment" ADD CONSTRAINT "CreditComment_creditId_fkey" FOREIGN KEY ("creditId") REFERENCES "Credit"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "Reviewer" ADD CONSTRAINT "Reviewer_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "DrawingRequirement" ADD CONSTRAINT "DrawingRequirement_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "ESDExcellenceOpportunity" ADD CONSTRAINT "ESDExcellenceOpportunity_projectId_fkey" FOREIGN KEY ("projectId") REFERENCES "Project"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "ESDExcellenceOpportunity" ADD CONSTRAINT "ESDExcellenceOpportunity_creditId_fkey" FOREIGN KEY ("creditId") REFERENCES "Credit"("id") ON DELETE SET NULL ON UPDATE CASCADE;
