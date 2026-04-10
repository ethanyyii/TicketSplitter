import { motion } from "framer-motion";
import UploadCard from "@/components/UploadCard";

/**
 * TicketSplitter 首頁 - 頂級視覺設計版本
 *
 * 設計理念：
 * - 極簡高對比風格（Apple 官網等級）
 * - 藍紫漸層文字效果
 * - 毛玻璃質感卡片設計
 * - 精細的拖曳互動動畫
 * - 無 emoji 的專業視覺
 */

const containerVariants = {
  hidden: { opacity: 0 },
  visible: {
    opacity: 1,
    transition: {
      staggerChildren: 0.1,
      delayChildren: 0.18,
    } as any,
  },
};

const itemVariants = {
  hidden: { opacity: 0, y: 20 },
  visible: {
    opacity: 1,
    y: 0,
    transition: { duration: 0.6 },
  },
};

export default function Home() {
  return (
    <div className="min-h-screen bg-white">
      {/* Hero Section */}
      <section className="hero-section">
        <div className="container">
          <motion.div
            className="hero-content"
            initial={{ opacity: 0, y: 30 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.8 } as any}
          >
            <h1 className="hero-title">
              <span className="block mb-2">機票費用</span>
              <span className="gradient-text">自動分帳</span>
            </h1>

            <p className="hero-subtitle">
              TicketSplitter 幫助您輕鬆上傳多家航空公司的機票，
              <br />
              自動計算分帳金額，告別複雜的手動計算。
            </p>
          </motion.div>
        </div>
      </section>

      {/* Bento Box 卡片區域 */}
      <section className="py-8 px-4 md:py-12">
        <div className="container">
          <motion.div
            className="bento-grid bento-grid-extended"
            variants={containerVariants}
            initial="hidden"
            animate="visible"
          >
            <motion.div variants={itemVariants}>
              <UploadCard
                airline="泰國獅子航空"
                airlineKey="thailionair"
                badgeCode="SL"
                description="上傳泰國獅子航空的機票，快速計算分帳"
              />
            </motion.div>

            <motion.div variants={itemVariants}>
              <UploadCard
                airline="台灣虎航"
                airlineKey="tigerair"
                badgeCode="IT"
                description="上傳台灣虎航的機票，輕鬆分帳"
              />
            </motion.div>

            <motion.div variants={itemVariants}>
              <UploadCard
                airline="酷航"
                airlineKey="scoot"
                badgeCode="TR"
                description="上傳酷航的機票，快速結算"
              />
            </motion.div>

            <motion.div variants={itemVariants}>
              <UploadCard
                airline="亞洲航空"
                airlineKey="airasia"
                badgeCode="AK"
                description="上傳亞洲航空（AirAsia）機票，快速拆分"
              />
            </motion.div>
          </motion.div>

          <motion.div
            className="text-center mt-12 md:mt-16"
            variants={itemVariants}
            initial="hidden"
            animate="visible"
          >
            <p className="text-sm text-slate-500 max-w-2xl mx-auto">
              支援 PDF、JPG 和 PNG 格式。您的機票資料完全安全，
              <br className="hidden md:block" />
              我們不會儲存任何個人資訊。
            </p>
          </motion.div>
        </div>
      </section>

      <div className="divider" />

      <section className="py-12 px-4">
        <div className="container text-center">
          <p className="text-xs text-slate-400">
            TicketSplitter © 2024. 機票分帳 SaaS 平台
          </p>
        </div>
      </section>
    </div>
  );
}
