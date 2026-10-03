import { Router, type IRouter } from "express";
import healthRouter from "./health";
import holidayMapRouter from "./holidayMap";
import rosterRouter from "./roster";
import scheduleOcrRouter from "./scheduleOcr";
import deviceAccessRouter from "./deviceAccess";

const router: IRouter = Router();

router.use(healthRouter);
router.use(holidayMapRouter);
router.use(rosterRouter);
router.use(scheduleOcrRouter);
router.use(deviceAccessRouter);

export default router;
