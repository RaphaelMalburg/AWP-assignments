"use client";

import dynamic from "next/dynamic";

const CandlestickChart = dynamic(
  () => import("./candlestick-chart-inner"),
  { ssr: false }
);

export default CandlestickChart;
