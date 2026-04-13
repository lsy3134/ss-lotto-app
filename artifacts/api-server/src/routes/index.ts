import { Router, type IRouter } from "express";
import healthRouter from "./health";
import holidayMapRouter from "./holidayMap";

const router: IRouter = Router();

router.use(healthRouter);
router.use(holidayMapRouter);

export default router;
