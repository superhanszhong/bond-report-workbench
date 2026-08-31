import type { Metadata } from "next";
import Workbench from "./workbench";

export const metadata: Metadata = {
  title: "利率债发行工作台",
  description: "地方债日表、利差图、四周滚动分析与周报生成的一体化工作台",
};

export default function Home() {
  return <Workbench />;
}
