import { Router, type IRouter } from "express";
import healthRouter from "./health";
import holidayMapRouter from "./holidayMap";
import rosterRouter from "./roster";

const router: IRouter = Router();

router.use(healthRouter);
router.use(holidayMapRouter);
router.use(rosterRouter);

export default router;
