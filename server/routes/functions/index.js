import { Router } from 'express';
import exportExcel from './exportExcel.js';
import generateCreditComment from './generateCreditComment.js';
import generateExcellence from './generateExcellence.js';
import parseBessPdf from './parseBessPdf.js';
import parseMatrixExcel from './parseMatrixExcel.js';
import regenerateContent from './regenerateContent.js';
import submitReview from './submitReview.js';

const router = Router();

router.post('/exportExcel', exportExcel);
router.post('/generateCreditComment', generateCreditComment);
router.post('/generateExcellence', generateExcellence);
router.post('/parseBessPdf', parseBessPdf);
router.post('/parseMatrixExcel', parseMatrixExcel);
router.post('/regenerateContent', regenerateContent);
router.post('/submitReview', submitReview);

export default router;
