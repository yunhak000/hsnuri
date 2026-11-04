import { formatKstDateTime } from "./datetime";

export const todayKst = () => formatKstDateTime().slice(0, 10);

export const makeDatedFileName = (suffix: string) => `${todayKst()}_${suffix}`;
