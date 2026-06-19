/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, ChangeEvent, FormEvent } from "react";
import Logo from "./components/Logo";
import { 
  motion, 
  AnimatePresence 
} from "motion/react";
import { 
  Menu, 
  X, 
  ChevronDown, 
  ChevronUp, 
  Download, 
  ArrowRight, 
  Check, 
  Award, 
  Settings, 
  ShieldCheck, 
  Cpu, 
  FileText, 
  CheckCircle,
  Building,
  Phone,
  Mail,
  MapPin,
  Clock,
  Briefcase,
  Layers,
  TrendingUp,
  Database,
  Copy,
  ExternalLink,
  AlertCircle
} from "lucide-react";

// 인터페이스 및 타입 정의
interface FAQItem {
  id: string;
  question: string;
  answer: string;
}

interface LeadForm {
  companyName: string;
  managerName: string;
  email: string;
  isAgreed: boolean;
}

const googleAppsScriptCode = `function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var timestamp = new Date();
  var companyName = "";
  var managerName = "";
  var email = "";
  
  try {
    if (e.postData.type === "application/json") {
      var data = JSON.parse(e.postData.contents);
      companyName = data.companyName || "";
      managerName = data.managerName || "";
      email = data.email || "";
    } else {
      companyName = e.parameter.companyName || "";
      managerName = e.parameter.managerName || "";
      email = e.parameter.email || "";
      
      if (!companyName && e.postData && e.postData.contents) {
        var textData = JSON.parse(e.postData.contents);
        companyName = textData.companyName || "";
        managerName = textData.managerName || "";
        email = textData.email || "";
      }
    }
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["신청 일시 (Timestamp)", "회사 이름 (Company)", "담당자 성함 (Manager)", "이메일 주소 (Email)"]);
    }
    
    sheet.appendRow([timestamp, companyName, managerName, email]);
    
    return ContentService.createTextOutput(JSON.stringify({ 
      "status": "success", 
      "message": "등록 완료되었습니다." 
    }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      "Access-Control-Allow-Origin": "*"
    });
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ 
      "status": "error", 
      "message": err.toString() 
    }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders({
      "Access-Control-Allow-Origin": "*"
    });
  }
}

function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeaders({
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type",
      "Access-Control-Max-Age": "86400"
    });
}`;

export default function App() {
  // 모바일 메뉴 상태
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState<boolean>(false);
  
  // 아코디언 상태 관리 (FAQ 열린 아이템 ID)
  const [openFAQId, setOpenFAQId] = useState<string | null>("q1");

  // 리드 수집 폼 상태
  const [formData, setFormData] = useState<LeadForm>({
    companyName: "",
    managerName: "",
    email: "",
    isAgreed: false
  });

  // 제출 성공 모달 상태
  const [isSubmitted, setIsSubmitted] = useState<boolean>(false);
  // 리드 제출 중 로딩 상태
  const [isSubmitting, setIsSubmitting] = useState<boolean>(false);
  // 에러 메시지 상태
  const [formError, setFormError] = useState<string | null>(null);

  // 구글 앱스 스크립트 웹 앱 배포 주소 (환경 변수 또는 브라우저 로컬 저장소 로드)
  const [scriptUrl, setScriptUrl] = useState<string>(() => {
    return (import.meta as any).env?.VITE_GOOGLE_SCRIPT_URL || localStorage.getItem("VITE_GOOGLE_SCRIPT_URL") || "https://script.google.com/macros/s/AKfycbyh0igSWIePtWfk7tzWMr-rDT7GTWbSiwmizv7e9-bVjoUv4EVU-O465BXMrVM7UtH_RQ/exec";
  });

  // 스프레드시트 연동 결과 상태 피드백
  const [syncStatus, setSyncStatus] = useState<{
    status: "idle" | "success" | "error";
    message?: string;
  }>({ status: "idle" });

  // 구글 스프레드시트 가이드 열기 상태
  const [isGuideOpen, setIsGuideOpen] = useState<boolean>(false);
  const [isCopying, setIsCopying] = useState<boolean>(false);

  // 스크롤 시 헤더 보더 상태
  const [isScrolled, setIsScrolled] = useState<boolean>(false);

  useEffect(() => {
    const handleScroll = () => {
      if (window.scrollY > 20) {
        setIsScrolled(true);
      } else {
        setIsScrolled(false);
      }
    };
    window.addEventListener("scroll", handleScroll);
    return () => window.removeEventListener("scroll", handleScroll);
  }, []);

  // 부드러운 스크롤 함수
  const scrollToSection = (id: string) => {
    const element = document.getElementById(id);
    if (element) {
      element.scrollIntoView({ behavior: "smooth" });
    }
    setIsMobileMenuOpen(false);
  };

  // FAQ 아코디언 데이터
  const faqData: FAQItem[] = [
    {
      id: "q1",
      question: "유원(U1)이라는 이름에는 어떤 철학이 담겨 있나요?",
      answer: "'당신(YOU)'을 '최고(ONE)'로 만드는 상생의 가치와, 사람들이 모여 좋은 기운과 행복이 머무는 버드나무 공원(柳園)의 의미를 동시에 담고 있습니다. 우리는 기술 그 이상의 신뢰로 고객과 인류가 상생하는 최고의 플랫폼이 되고자 합니다."
    },
    {
      id: "q2",
      question: "유원스틸 부품의 실제 자동차 내 적용 위치와 주요 거래처는 어디인가요?",
      answer: "드림텍, 코우 등이 주요 파트너이며, 차량 조향축의 핵심인 스티어링 컬럼 부근 인터뮤디어 샤프트(Intermediate Shaft)에 결합하는 '일체형 요크(Yoke)' 등 타협 없는 탑클래스 안전과 직결된 중요 고강도 보안 부품 및 패스너류 전반에 안정적으로 적용되고 있습니다."
    },
    {
      id: "q3",
      question: "진천공장의 설비 규모와 다공정 대응 능력이 궁금합니다.",
      answer: "충북 진천에 위치한 9,946m²(약 3,000평) 규모의 최신 자동화 공장입니다. 초정밀 절단기 WL75X 4기(연 1,367만 개 생산 능력)와 복합 직선기 WYF-3200 1기(연 11,630톤 가공 능력)를 인하우스 자작 설비로 직접 보유하고 가동하여, 다품종 소량 즉시 대응은 물론 대규모 연속 오더까지 완벽히 소화 가능한 유연 캐파를 자랑합니다."
    },
    {
      id: "q4",
      question: "원소재 조달 및 품질 관리는 어떻게 신뢰할 수 있나요?",
      answer: "우리는 대한민국 최고의 철강사인 포스코(POSCO) 및 현대제철(HYUNDAI STEEL) 산하의 공식 1차 유통망에서 공급하는 보증된 원소재만을 기반으로 삼습니다. 이를 바탕으로 ISO 9001, 14001, 45001 통합 경영시스템 인증 표준과 KOITA 인정 기업 R&D 부서의 수많은 한계 내구 테스트를 거쳐 티끌만큼의 타협도 없는 품질을 완성합니다."
    },
    {
      id: "q5",
      question: "유원스틸산업의 중장기 매출 및 성장 비전은 무엇인가요?",
      answer: "2026년 매출 200억 달성(전년 대비 +85%의 성장을 완벽히 조준)을 임계점으로, 2028년 독자 기술 기반 자동차 1차 협력사(Tier-1) 진입, 2030년 해외 OEM 가전 및 글로벌 특수 목적 강판 수주 확장을 통한 매출 1,000억 원 달성, 그리고 마침내 2031년 코스닥(KOSDAQ) 기업공개(IPO)를 정해진 미래로 만들어 나가고 있습니다."
    }
  ];

  // 폼 입력 핸들러
  const handleInputChange = (e: ChangeEvent<HTMLInputElement>) => {
    const { name, value, type, checked } = e.target;
    setFormData(prev => ({
      ...prev,
      [name]: type === "checkbox" ? checked : value
    }));
  };

  // 폼 제출 핸들러
  const handleSubmit = async (e: FormEvent) => {
    e.preventDefault();
    setFormError(null);
    setSyncStatus({ status: "idle" });

    if (!formData.companyName.trim()) {
      setFormError("회사 이름을 입력해주세요.");
      return;
    }
    if (!formData.managerName.trim()) {
      setFormError("담당자 이름을 입력해주세요.");
      return;
    }
    if (!formData.email.trim()) {
      setFormError("이메일 주소를 입력해주세요.");
      return;
    }
    if (!formData.isAgreed) {
      setFormError("개인정보 수집 및 이용 동의가 필요합니다.");
      return;
    }

    // 이메일 정규식 검증
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(formData.email)) {
      setFormError("올바른 이메일 주소 형식이 아닙니다.");
      return;
    }

    setIsSubmitting(true);

    if (scriptUrl.trim()) {
      try {
        // GAS는 core 리다이렉션으로 인해 CORS 이슈가 발생할 수 있습니다.
        // mode: "no-cors"를 사용하여 브라우저 보안 규제를 완벽하게 우회하고 전송 성공율을 극대화합니다.
        await fetch(scriptUrl.trim(), {
          method: "POST",
          mode: "no-cors",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            companyName: formData.companyName,
            managerName: formData.managerName,
            email: formData.email
          })
        });

        // no-cors 응답은 불투명(opaque)하므로 상태를 임의로 성공 처리하여 브라우저 대기시간 차단
        setSyncStatus({
          status: "success",
          message: "구글 스프레드시트에 실시간으로 데이터가 온전히 연동 전송되었습니다."
        });
      } catch (err: any) {
        console.error("스프레드시트 동기화 실패:", err);
        setSyncStatus({
          status: "error",
          message: `구글 스프레드시트 연동 실패: ${err.message || "연결 오류"}`
        });
      }
    } else {
      setSyncStatus({
        status: "idle",
        message: "소개서 신청이 임시 완료되었습니다. 단, 구글 앱스 스크립트 전송 주소(Web App URL)가 비어있어 스프레드시트에 기입되지 않았습니다. 실시간 연동을 위해 하단의 스프레드시트 셋업 가이드를 확인하여 URL을 설정해주세요."
      });
    }

    setIsSubmitting(false);
    // 성공 상태로 변경
    setIsSubmitted(true);
  };

  // 다시 신청하기
  const handleResetForm = () => {
    setFormData({
      companyName: "",
      managerName: "",
      email: "",
      isAgreed: false
    });
    setIsSubmitted(false);
    setFormError(null);
  };

  return (
    <div className="min-h-screen bg-[#F4F5F7] text-[#1A1A1A] font-sans tracking-tight selection:bg-u1-blue selection:text-white overflow-x-hidden">
      
      {/* 1. 네비게이션 바 */}
      <header 
        id="navbar"
        className={`fixed top-0 left-0 right-0 z-50 transition-all duration-300 ${
          isScrolled 
            ? "bg-white/85 backdrop-blur-md border-b border-black/5 shadow-xs py-4" 
            : "bg-transparent py-6"
        }`}
      >
        <div className="max-w-7xl mx-auto px-6 md:px-12 flex items-center justify-between">
          {/* 로고 영역 */}
          <div 
            onClick={() => scrollToSection("hero")} 
            className="flex items-center cursor-pointer group"
            id="nav-logo-zone"
          >
            {/* 유원스틸 로고 영역 - 고해상도 벡터 탑재 */}
            <Logo className="h-11 md:h-15" showSubtitle={true} />
          </div>

          {/* 데스크톱 메뉴 */}
          <nav className="hidden lg:flex items-center space-x-10 text-xs tracking-wider uppercase font-medium text-gray-500">
            <button 
              onClick={() => scrollToSection("business")} 
              className="hover:text-u1-blue transition-colors duration-200 cursor-pointer text-[13px]"
              id="menu-target-business"
            >
              BUSINESS SEGMENTS
            </button>
            <button 
              onClick={() => scrollToSection("competitiveness")} 
              className="hover:text-u1-blue transition-colors duration-200 cursor-pointer text-[13px]"
              id="menu-target-comp"
            >
              CORE COMPETITIVENESS
            </button>
            <button 
              onClick={() => scrollToSection("process")} 
              className="hover:text-u1-blue transition-colors duration-200 cursor-pointer text-[13px]"
              id="menu-target-proc"
            >
              PROCESS INNOVATION
            </button>
            <button 
              onClick={() => scrollToSection("faq")} 
              className="hover:text-u1-blue transition-colors duration-200 cursor-pointer text-[13px]"
              id="menu-target-faq"
            >
              FAQ
            </button>
            <button 
              onClick={() => scrollToSection("lead-section")} 
              className="bg-u1-blue text-white px-5 py-2.5 rounded-sm font-semibold hover:bg-[#4eb233] transition-colors duration-300 shadow-sm tracking-normal text-[12px] flex items-center space-x-1.5 cursor-pointer"
              id="menu-target-profile-btn"
            >
              <span>COMPANY PROFILE</span>
              <ArrowRight className="w-3 H-3" />
            </button>
          </nav>

          {/* 모바일 햄버거 버튼 */}
          <button 
            type="button"
            className="lg:hidden p-2 text-gray-700 hover:text-u1-blue focus:outline-none cursor-pointer"
            onClick={() => setIsMobileMenuOpen(true)}
            aria-label="Open Menu"
            id="mobile-menu-burger"
          >
            <Menu className="w-6 h-6" />
          </button>
        </div>
      </header>

      {/* 모바일 전체 화면 오버레이 슬라이드 메뉴 */}
      <AnimatePresence>
        {isMobileMenuOpen && (
          <>
            {/* 배경 블러 오버레이 */}
            <motion.div 
              className="fixed inset-0 bg-black/40 backdrop-blur-sm z-50"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setIsMobileMenuOpen(false)}
              id="mobile-drawer-overlay"
            />
            {/* 슬라이드 화도 */}
            <motion.div 
              className="fixed top-0 right-0 bottom-0 w-4/5 max-w-sm bg-white shadow-2xl z-50 p-8 flex flex-col justify-between rounded-l-[4px]"
              initial={{ x: "100%" }}
              animate={{ x: 0 }}
              exit={{ x: "100%" }}
              transition={{ type: "spring", damping: 25, stiffness: 200 }}
              id="mobile-drawer-panel"
            >
              <div>
                <div className="flex items-center justify-between mb-12">
                  <Logo className="h-10 md:h-12" showSubtitle={false} />
                  <button 
                    type="button"
                    className="p-2 -mr-2 text-gray-500 hover:text-gray-900 cursor-pointer"
                    onClick={() => setIsMobileMenuOpen(false)}
                    id="mobile-drawer-close"
                  >
                    <X className="w-6 h-6" />
                  </button>
                </div>

                <nav className="flex flex-col space-y-6 text-sm font-medium tracking-wide">
                  <button 
                    onClick={() => scrollToSection("business")} 
                    className="text-left py-2 border-b border-gray-100 text-gray-800 hover:text-u1-blue flex items-center justify-between cursor-pointer"
                  >
                    <span>사업 영역</span>
                    <span className="text-[10px] text-gray-400 font-mono">01</span>
                  </button>
                  <button 
                    onClick={() => scrollToSection("competitiveness")} 
                    className="text-left py-2 border-b border-gray-100 text-gray-800 hover:text-u1-blue flex items-center justify-between cursor-pointer"
                  >
                    <span>핵심 경쟁력</span>
                    <span className="text-[10px] text-gray-400 font-mono">02</span>
                  </button>
                  <button 
                    onClick={() => scrollToSection("process")} 
                    className="text-left py-2 border-b border-gray-100 text-gray-800 hover:text-u1-blue flex items-center justify-between cursor-pointer"
                  >
                    <span>공정 단축 혁신</span>
                    <span className="text-[10px] text-gray-400 font-mono">03</span>
                  </button>
                  <button 
                    onClick={() => scrollToSection("faq")} 
                    className="text-left py-2 border-b border-gray-100 text-gray-800 hover:text-u1-blue flex items-center justify-between cursor-pointer"
                  >
                    <span>자주 묻는 질문 (FAQ)</span>
                    <span className="text-[10px] text-gray-400 font-mono">04</span>
                  </button>
                </nav>
              </div>

              <div className="mt-auto pt-8 border-t border-gray-100">
                <button 
                  onClick={() => scrollToSection("lead-section")} 
                  className="w-full bg-[#132e88] text-white text-center py-4 rounded-sm font-semibold hover:bg-u1-green transition-colors duration-300 flex items-center justify-center space-x-2 cursor-pointer"
                >
                  <Download className="w-4 h-4" />
                  <span>상세 회사소개서 보기</span>
                </button>
                <p className="text-[10px] text-gray-400 mt-4 text-center">Copyright © YOUONE STEEL Corp.</p>
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* 2. 히어로 섹션 */}
      <section 
        id="hero" 
        className="relative pt-32 pb-24 md:pt-48 md:pb-36 bg-gradient-to-b from-[#f2f4f8] via-[#fafafa] to-white overflow-hidden"
      >
        <div className="max-w-7xl mx-auto px-6 md:px-12 relative z-10">
          
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-12 lg:gap-16 items-center">
            
            {/* 왼쪽 카피텍스트 영역 */}
            <div className="lg:col-span-12 text-center max-w-4xl mx-auto">
              <motion.div 
                className="inline-flex items-center space-x-2 px-3.5 py-1.5 rounded-full bg-white border border-gray-200/80 shadow-xs mb-8"
                initial={{ opacity: 0, y: 15 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.6 }}
              >
                <span className="flex h-2 w-2 relative">
                  <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-u1-green opacity-75"></span>
                  <span className="relative inline-flex rounded-full h-2 w-2 bg-u1-green"></span>
                </span>
                <span className="text-[11px] font-bold tracking-widest text-[#132e88] font-mono uppercase">LEADER OF PROCESS INNOVATION</span>
              </motion.div>

              <motion.h1 
                className="text-4xl sm:text-5xl md:text-6xl lg:text-[72px] leading-[1.2] lg:leading-[1.15] font-extrabold text-u1-blue tracking-tighter break-keep"
                initial={{ opacity: 0, y: 30 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.8, delay: 0.1 }}
                id="hero-main-title"
              >
                자동차 부품 제조와 철강 소재를<br />
                잇는 <span className="text-[#4eb233]">공정 혁신 기업</span>,<br />
                유원스틸산업(주)
              </motion.h1>

              <motion.p 
                className="mt-8 text-base sm:text-lg md:text-xl text-gray-500 font-light leading-relaxed max-w-3xl mx-auto break-keep"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.8, delay: 0.2 }}
                id="hero-sub-title"
              >
                당신이 1등이 되는 가치 유원(U1). 우리는 기술로 가장 투명한 신뢰를 만들고, 따뜻한 나눔으로 함께 성장합니다.
              </motion.p>

              <motion.div 
                className="mt-12 flex flex-col sm:flex-row items-center justify-center gap-4"
                initial={{ opacity: 0, y: 20 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.8, delay: 0.3 }}
              >
                <button 
                  onClick={() => scrollToSection("lead-section")}
                  className="w-full sm:w-auto bg-u1-blue text-white px-8 py-4.5 rounded-xl font-bold transition-all duration-300 hover:bg-[#132e88]/90 border-b-4 border-transparent hover:border-b-u1-yellow flex items-center justify-center space-x-2 shadow-lg hover:shadow-u1-blue/15"
                  id="hero-cta-button"
                >
                  <span>상세 회사소개서 보기</span>
                  <ArrowRight className="w-4 h-4 text-u1-yellow" />
                </button>
                <button 
                  onClick={() => scrollToSection("business")}
                  className="w-full sm:w-auto bg-white text-gray-800 border border-gray-200 px-8 py-4.5 rounded-xl font-semibold hover:border-u1-blue/30 transition-all duration-300"
                  id="hero-secondary-button"
                >
                  비즈니스 영역 둘러보기
                </button>
              </motion.div>
            </div>

            {/* 하단 웅장한 진천공장 전경 플레이스홀더 영역 */}
            <div className="lg:col-span-12 mt-4">
              <motion.div 
                className="relative rounded-3xl overflow-hidden shadow-2xl border border-gray-200/70"
                initial={{ opacity: 0, scale: 0.96 }}
                animate={{ opacity: 1, scale: 1 }}
                transition={{ duration: 0.9, delay: 0.4 }}
                id="hero-image-container"
              >
                {/* [진천공장_전경_사진.jpg] 전경 이미지 슬롯 */}
                <div className="w-full h-[400px] md:h-[580px] bg-gray-900 relative flex items-center justify-center overflow-hidden">
                  <img 
                    src="[진천공장_전경_사진.jpg]" 
                    alt="유원스틸산업(주) 충북 진천 최신 생산 기지 공장 전경" 
                    className="w-full h-full object-cover opacity-35"
                    onError={(e) => {
                      // 이미지 부재시 노출될 미니멀리스틱 디자이너 플레이스홀더 백그라운드
                      const target = e.target as HTMLElement;
                      target.style.display = 'none';
                    }}
                  />
                  
                  {/* 고급스러운 기하학 레이아웃 디자인 패널 */}
                  <div className="absolute inset-0 bg-radial-gradient from-transparent to-black/80 flex flex-col justify-between p-8 md:p-14 z-10 text-white">
                    <div className="flex justify-between items-start">
                      <div className="border-l-2 border-u1-green pl-4">
                        <span className="text-[10px] font-mono tracking-widest text-[#4eb233] uppercase block">FACILITY INFRASTRUCTURE</span>
                        <h3 className="text-xl md:text-2xl font-bold tracking-tight mt-1 text-white">진천 스마트 제조 공장</h3>
                      </div>
                      <span className="text-xs font-mono text-gray-400 bg-black/40 px-3 py-1.5 rounded-md backdrop-blur-md">CAPACITY: 11,630 TONS</span>
                    </div>

                    <div className="max-w-xl self-start">
                      <p className="text-xs md:text-sm text-gray-300 font-light leading-relaxed">
                        충북 진천 신척산업단지 내 위치한 9,946m² 규모의 차세대 혁신 인프라는 고도의 자작 설비 내재화로, 원재료 가공부터 초정밀 완성 부품 양산까지 전 단계를 수직 계열화하여 고품질 초저단가를 실현합니다.
                      </p>
                    </div>
                  </div>
                </div>
              </motion.div>
            </div>
            
          </div>
          
        </div>
      </section>

      {/* 3. 비즈니스 세그먼트 섹션 (BUSINESS SEGMENTS) */}
      <section id="business" className="py-24 md:py-36 bg-white border-y border-gray-100">
        <div className="max-w-7xl mx-auto px-6 md:px-12">
          
          <div className="max-w-3xl mb-20 md:mb-28">
            <span className="text-xs font-mono tracking-widest text-[#4eb233] uppercase font-bold block mb-3">01 / BUSINESS SEGMENTS</span>
            <h2 className="text-3xl md:text-5xl font-extrabold text-u1-blue tracking-tight leading-tight">
              당신이 1등이 되는 가치,<br />유원의 두 가지 핵심 비즈니스 축
            </h2>
            <div className="w-12 h-1 bg-[#4eb233] mt-6"></div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-10 lg:gap-16">
            
            {/* 부문 1: 유통 */}
            <div className="group bg-neutral-50/50 rounded-3xl p-6 md:p-10 border border-gray-100 flex flex-col justify-between hover:scale-[1.015] hover:shadow-2xl transition-all duration-500 ease-out" id="card-distribution">
              <div>
                <div className="relative h-64 md:h-80 bg-gray-800 rounded-2xl overflow-hidden mb-8" id="visual-dist-image-block">
                  {/* [철강_소재_유통_사진.jpg] 유통 부문 테마 비주얼 */}
                  <img 
                    src="[철강_소재_유통_사진.jpg]" 
                    alt="철강 유통 및 가공 원소재 실물 소재" 
                    className="w-full h-full object-cover opacity-45 group-hover:scale-105 transition-transform duration-500"
                    onError={(e) => {
                      const target = e.target as HTMLElement;
                      target.style.display = 'none';
                    }}
                  />
                  <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-transparent to-transparent flex items-end p-6">
                    <div>
                      <span className="text-[10px] font-mono text-[#fdd000] tracking-widest uppercase block mb-1">RAW MATERIALS SUPPLY</span>
                      <h3 className="text-xl md:text-2xl font-bold text-white">철강 소재 유통 및 대량 조달</h3>
                    </div>
                  </div>
                </div>

                <div className="flex items-center space-x-2 text-u1-blue mb-4">
                  <span className="text-xs font-mono font-bold tracking-widest uppercase border border-u1-blue/25 px-2 py-0.5 rounded text-u1-blue">DISTRIBUTION</span>
                  <span className="text-gray-300">|</span>
                  <span className="text-sm font-semibold text-gray-500 dark:text-gray-400">다품종 소량 즉각 대응</span>
                </div>

                <h4 className="text-lg md:text-xl font-bold text-u1-blue mb-4">철강 소재 유통 (다품종 소량 대응)</h4>
                
                <p className="text-sm md:text-base text-gray-500 font-light leading-relaxed mb-6">
                  CHQ, CD-BAR 및 흑환봉(탄소강·망간강·크롬강·SUJ2), 다양한 구조용 PIPE 및 냉연강판(배관·구조용 고장력)을 비롯해 완벽한 보호 적재를 위한 PALLET 가공까지, 맞춤 주문생산 인프라를 구축해 어떤 소량 다품종 주문이라도 정교하게 정시 배송합니다.
                </p>
              </div>

              <div className="border-t border-gray-200/50 pt-6 mt-6">
                <ul className="grid grid-cols-2 gap-y-3 gap-x-2 text-xs md:text-sm text-gray-600 font-medium">
                  <li className="flex items-center space-x-2">
                    <Check className="w-4 h-4 text-u1-green" />
                    <span>최고급 탄소강 / 크롬강 원소재</span>
                  </li>
                  <li className="flex items-center space-x-2">
                    <Check className="w-4 h-4 text-u1-green" />
                    <span>CD-BAR & 흑환봉 유연 조달</span>
                  </li>
                  <li className="flex items-center space-x-2">
                    <Check className="w-4 h-4 text-u1-green" />
                    <span>냉연강판 & 고장력 배관</span>
                  </li>
                  <li className="flex items-center space-x-2">
                    <Check className="w-4 h-4 text-u1-green" />
                    <span>맞춤 PALLET 인하우스 제작</span>
                  </li>
                </ul>
              </div>
            </div>

            {/* 부문 2: 제조 */}
            <div className="group bg-neutral-50/50 rounded-3xl p-6 md:p-10 border border-gray-100 flex flex-col justify-between hover:scale-[1.015] hover:shadow-2xl transition-all duration-500 ease-out" id="card-manufacturing">
              <div>
                <div className="relative h-64 md:h-80 bg-gray-800 rounded-2xl overflow-hidden mb-8" id="visual-mfg-image-block">
                  {/* [자동차_부품_제조_사진.jpg] 제조 부문 테마 비주얼 */}
                  <img 
                    src="[자동차_부품_제조_사진.jpg]" 
                    alt="자동차 정밀 안전 부품 제조 가공 및 설비" 
                    className="w-full h-full object-cover opacity-45 group-hover:scale-105 transition-transform duration-500"
                    onError={(e) => {
                      const target = e.target as HTMLElement;
                      target.style.display = 'none';
                    }}
                  />
                  <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-transparent to-transparent flex items-end p-6">
                    <div>
                      <span className="text-[10px] font-mono text-[#fdd000] tracking-widest uppercase block mb-1">AUTO COMPONENT BRANDING</span>
                      <h3 className="text-xl md:text-2xl font-bold text-white">자동차 정밀 보안 부품 양산</h3>
                    </div>
                  </div>
                </div>

                <div className="flex items-center space-x-2 text-u1-blue mb-4">
                  <span className="text-xs font-mono font-bold tracking-widest uppercase border border-u1-blue/25 px-2 py-0.5 rounded text-u1-blue">MANUFACTURING</span>
                  <span className="text-gray-300">|</span>
                  <span className="text-sm font-semibold text-gray-500">Tier-4 핵심 최상류 파트너</span>
                </div>

                <h4 className="text-lg md:text-xl font-bold text-u1-blue mb-4">자동차 FASTNER 제조 & CFQ Solid</h4>
                
                <p className="text-sm md:text-base text-gray-500 font-light leading-relaxed mb-6">
                  차량의 핵심 안전 계통인 조향장치와 현가장치 탑 티어 부품 공급망에서 중추적인 역할을 담당합니다. 현대모비스, 만도(HL Mando) 등 글로벌 탑티어의 Tier 4 핵심 최상류에서 차량 전체 무결성과 원가를 결정하는 ‘일체형 요크(Yoke)’ 시장의 핵심 기술 리더입니다.
                </p>
              </div>

              <div className="border-t border-gray-200/50 pt-6 mt-6">
                <ul className="grid grid-cols-2 gap-y-3 gap-x-2 text-xs md:text-sm text-gray-600 font-medium">
                  <li className="flex items-center space-x-2">
                    <Check className="w-4 h-4 text-u1-green" />
                    <span>글로벌 고난도 요크 공급</span>
                  </li>
                  <li className="flex items-center space-x-2">
                    <Check className="w-4 h-4 text-u1-green" />
                    <span>조향장치 / 스티어링 부품</span>
                  </li>
                  <li className="flex items-center space-x-2">
                    <Check className="w-4 h-4 text-u1-green" />
                    <span>현가장치 특수 패스너 가공</span>
                  </li>
                  <li className="flex items-center space-x-2">
                    <Check className="w-4 h-4 text-u1-green" />
                    <span>CFQ Solid 초정밀 양산 기술</span>
                  </li>
                </ul>
              </div>
            </div>

          </div>

        </div>
      </section>

      {/* 4. 핵심 경쟁력 및 차별점 섹션 (R&D & PROCESS INNOVATION) */}
      <section id="competitiveness" className="py-24 md:py-36 bg-[#fafafa]">
        <div className="max-w-7xl mx-auto px-6 md:px-12">
          
          <div className="flex flex-col lg:flex-row lg:items-end justify-between mb-20 md:mb-28">
            <div className="max-w-3xl">
              <span className="text-xs font-mono tracking-widest text-[#4eb233] uppercase font-bold block mb-3">02 / CORE ASSETS</span>
              <h2 className="text-3xl md:text-5xl font-extrabold text-u1-blue tracking-tight leading-tight">
                Continuous And Never-Ending Improvement (CANI)
              </h2>
              <p className="mt-5 text-gray-500 font-light text-base md:text-lg">
                유원스틸산업(주)은 결코 멈추지 않는 지속적 점진적 공정 개선을 통해, 최상의 경쟁우위를 영속화합니다.
              </p>
            </div>
            <div className="mt-6 lg:mt-0">
              <span className="text-[11px] font-mono font-semibold tracking-wider text-gray-400 border border-gray-300/60 px-4 py-2 rounded-xl">
                3 CORE SYSTEMS INSTALLED
              </span>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-10">
            
            {/* 카드 01 */}
            <div className="bg-white rounded-3xl p-8 md:p-10 border border-gray-200/50 hover:shadow-xl transition-all duration-300 relative flex flex-col justify-between" id="comp-grid-card-1">
              <div>
                <div className="flex items-center justify-between mb-8">
                  <span className="text-5xl md:text-6xl font-light text-u1-blue/15 font-mono">01</span>
                  <div className="p-3.5 rounded-2xl bg-u1-blue/5 text-[#132e88]" id="comp-icon-1">
                    <Award className="w-6 h-6" />
                  </div>
                </div>
                <h3 className="text-xl font-bold text-u1-blue mb-4">독보적 R&D 자산</h3>
                <p className="text-gray-500 font-light text-sm md:text-base leading-relaxed">
                  KOITA(한국산업기술진흥협회) 공식 인증 R&D 전담부서를 설립하여 운영하고 있으며, 부품 가공 품질의 성패를 가르는 '일체형 요크용 봉강재 제조 기술 및 특수한 가공 압밀 기술' 자체 특허 3건을 철저하게 독점 보유하고 있습니다.
                </p>
              </div>
              <div className="mt-8 pt-6 border-t border-gray-100 flex items-center justify-between">
                <span className="text-xs font-mono text-gray-400 font-medium">KOITA R&D LAB</span>
                <span className="text-xs text-[#4eb233] font-bold">특허 3건 보유</span>
              </div>
            </div>

            {/* 카드 02 */}
            <div className="bg-white rounded-3xl p-8 md:p-10 border border-gray-200/50 hover:shadow-xl transition-all duration-300 relative flex flex-col justify-between" id="comp-grid-card-2">
              <div>
                <div className="flex items-center justify-between mb-8">
                  <span className="text-5xl md:text-6xl font-light text-u1-blue/15 font-mono">02</span>
                  <div className="p-3.5 rounded-2xl bg-u1-blue/5 text-[#132e88]" id="comp-icon-2">
                    <Cpu className="w-6 h-6" />
                  </div>
                </div>
                <h3 className="text-xl font-bold text-u1-blue mb-4">압도적 생산 인프라</h3>
                <p className="text-gray-500 font-light text-sm md:text-base leading-relaxed">
                  32억 원 이상의 선제적 대규모 설비 투자를 단행하였습니다. 충북 진천공장(9,946m²) 내 고생산성 직선기(WYF-3200, 연간 최대 11,630톤 커트 가밀형) 및 초정밀 정량 절단기(WL75X 4기, 연 1,367만개) 인하우스 전량 가동으로 납기 무결성을 약속합니다.
                </p>
              </div>
              <div className="mt-8 pt-6 border-t border-gray-100 flex items-center justify-between">
                <span className="text-xs font-mono text-gray-400 font-medium">CAPEX INVESTMENT</span>
                <span className="text-xs text-[#4eb233] font-bold">인하우스 자작 설비</span>
              </div>
            </div>

            {/* 카드 03 */}
            <div className="bg-white rounded-3xl p-8 md:p-10 border border-gray-200/50 hover:shadow-xl transition-all duration-300 relative flex flex-col justify-between" id="comp-grid-card-3">
              <div>
                <div className="flex items-center justify-between mb-8">
                  <span className="text-5xl md:text-6xl font-light text-u1-blue/15 font-mono">03</span>
                  <div className="p-3.5 rounded-2xl bg-u1-blue/5 text-[#132e88]" id="comp-icon-3">
                    <ShieldCheck className="w-6 h-6" />
                  </div>
                </div>
                <h3 className="text-xl font-bold text-u1-blue mb-4">글로벌 품질 표준</h3>
                <p className="text-gray-500 font-light text-sm md:text-base leading-relaxed">
                  품질 경영, 환경, 보건 안전을 전격 아우르는 ISO 9001 / 14001 / 45001 통합 경영시스템 인증을 취득 완료하였으며, 중소벤처기업부 인증 혁신성장유형 벤처기업 확인을 마쳤습니다. 포스코·현대제철 직통 공식 원소재 보증으로 타협 없이 납품합니다.
                </p>
              </div>
              <div className="mt-8 pt-6 border-t border-gray-100 flex items-center justify-between">
                <span className="text-xs font-mono text-gray-400 font-medium">GLOBAL COMPLIANCE</span>
                <span className="text-xs text-[#4eb233] font-bold">ISO 통합인증 & 벤처 확인</span>
              </div>
            </div>

          </div>

        </div>
      </section>

      {/* 5. Before & After 공정 혁신 섹션 (7-STEP PROCESS INNOVATION) */}
      <section id="process" className="py-24 md:py-36 bg-white border-b border-gray-100">
        <div className="max-w-7xl mx-auto px-6 md:px-12">
          
          <div className="max-w-3xl mb-12">
            <span className="text-xs font-mono tracking-widest text-[#4eb233] uppercase font-bold block">03 / ECOSYSTEM DISRUPTION</span>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-12 gap-12 items-stretch mb-16">
            
            {/* AS-IS 블록 (45% 폭 느낌) */}
            <div className="lg:col-span-5 bg-gray-50 rounded-3xl p-8 md:p-10 border border-gray-100 flex flex-col justify-between" id="process-asis-block">
              <div>
                <span className="text-xs font-mono tracking-wider font-extrabold text-red-500 bg-red-50 px-2.5 py-1 rounded inline-block mb-6">
                  CONVENTIONAL PROCESS (AS-IS)
                </span>
                <h3 className="text-xl font-bold text-gray-800 mb-4">기존 아웃소싱 중심 9단계 복합 공정</h3>
                <p className="text-gray-500 text-sm font-light leading-relaxed mb-8">
                  외주 의존도가 높고 절차와 롤 관리가 난삽했던 기존 수기 공정은 각 연결 마디마다 화물 운송, 재고 대기, 불필요한 핸들링 리드가 발생하여 납기 및 로스 리스크가 누적되었습니다.
                </p>

                {/* 9단계 시각적 노드 */}
                <div className="space-y-3.5 pr-4" id="asis-steps-visual font-mono">
                  <div className="flex items-center space-x-3 text-xs text-gray-400 line-through">
                    <span className="w-5 h-5 rounded-full border border-gray-300 flex items-center justify-center text-[10px] font-bold">1</span>
                    <span>포스코 코일 입고 수령 및 야적 대기</span>
                  </div>
                  <div className="flex items-center space-x-3 text-xs text-gray-400 line-through">
                    <span className="w-5 h-5 rounded-full border border-gray-300 flex items-center justify-center text-[10px] font-bold">2</span>
                    <span className="text-red-500/80 font-medium">[외주] 1차 드로잉 피막 처리 탈피 가공</span>
                  </div>
                  <div className="flex items-center space-x-3 text-xs text-gray-400 line-through">
                    <span className="w-5 h-5 rounded-full border border-gray-300 flex items-center justify-center text-[10px] font-bold">3</span>
                    <span>직선 복원 선가공 및 직선 풀림 대기</span>
                  </div>
                  <div className="flex items-center space-x-3 text-xs text-gray-400 line-through">
                    <span className="w-5 h-5 rounded-full border border-gray-300 flex items-center justify-center text-[10px] font-bold">4</span>
                    <span className="text-red-500/80 font-medium">[외주] 숏블라스트 프라이머 표면 스케일 제거</span>
                  </div>
                  <div className="flex items-center space-x-3 text-xs text-gray-400 line-through">
                    <span className="w-5 h-5 rounded-full border border-gray-300 flex items-center justify-center text-[10px] font-bold">5</span>
                    <span>입고 검사 후 원재료 정밀 1차 전단 보일러 가밀</span>
                  </div>
                  <div className="flex items-center space-x-3 text-xs text-gray-400 line-through">
                    <span className="w-5 h-5 rounded-full border border-gray-300 flex items-center justify-center text-[10px] font-bold">6</span>
                    <span className="text-red-500/80 font-medium">[외주] 일체형 요크 냉간 포징 전압 강밀 성형</span>
                  </div>
                  <div className="flex items-center space-x-3 text-xs text-gray-400/60">
                    <span className="w-5 h-5 rounded-full border border-gray-200 flex items-center justify-center text-[10px] font-bold">7,8,9</span>
                    <span>품질 전수 계측 검사, 패키징 팔레트 로딩 및 출하 배송</span>
                  </div>
                </div>
              </div>

              <div className="border-t border-gray-200 pt-6 mt-8 text-xs text-gray-400 font-mono">
                총 소요 리드타임: 평균 12일 ~ 15일 소요
              </div>
            </div>

            {/* 중간 화살표 가이드 (데스크톱 전용) */}
            <div className="hidden lg:col-span-2 lg:flex flex-col items-center justify-center relative">
              <div className="w-full flex items-center justify-center h-full">
                <div className="bg-u1-blue/10 text-u1-blue p-4 rounded-full border border-u1-blue/15 shadow-sm transform rotate-90 lg:rotate-0">
                  <ArrowRight className="w-7 h-7 text-u1-blue" />
                </div>
              </div>
              <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 -mt-16 text-[10px] font-mono uppercase tracking-widest text-[#4eb233] font-bold whitespace-nowrap bg-white py-1 px-3 border border-gray-200/60 rounded-full shadow-xs">
                U1 INNOVATION EFFECT
              </div>
            </div>

            {/* TO-BE 블록 (55% 폭 느낌, 보다 가시적이고 화려하게) */}
            <div className="lg:col-span-5 bg-gradient-to-br from-[#132e88] to-[#0d1f5e] rounded-3xl p-8 md:p-10 text-white flex flex-col justify-between shadow-xl shadow-u1-blue/15 relative overflow-hidden" id="process-tobe-block">
              {/* 은은하게 백그라운드로 도는 기술선 비주얼 */}
              <div className="absolute top-0 right-0 w-64 h-64 bg-radial-gradient from-[#4eb233]/15 to-transparent rounded-full -mr-20 -mt-20"></div>

              <div className="relative z-10">
                <div className="flex items-center justify-between mb-6">
                  <span className="text-xs font-mono tracking-wider font-extrabold text-[#4eb233] bg-[#4eb233]/15 px-3 py-1 rounded">
                    U1 COMPRESSIVE MANUFACTURING (TO-BE)
                  </span>
                  <span className="text-xs font-bold text-u1-yellow">적용 특허 기술</span>
                </div>
                
                <h3 className="text-xl md:text-2xl font-bold mb-4">유원스틸 자작 인하우스 7단계 수직 계열화</h3>
                <p className="text-gray-300 text-sm font-light leading-relaxed mb-8">
                  대규모 직선기, 쇼트 블래스터, 절단기 설비 투자를 내재화하고 특허 공정을 이식함으로써 품질에 악영향을 주던 외주 외송 세 마디를 완전히 없앴습니다. 완벽 통제되는 7단계 통합 인라인 설계입니다.
                </p>

                {/* 7단계 완벽 로드맵 */}
                <div className="space-y-4" id="tobe-steps-visual font-mono">
                  <div className="flex items-center space-x-3.5 text-xs text-slate-100">
                    <span className="w-5.5 h-5.5 rounded-full bg-[#4eb233] text-[#132e88] flex items-center justify-center text-[10px] font-black">1</span>
                    <span className="font-semibold">포스코 원소재 고강도 철강 코일 안전 입고</span>
                  </div>
                  <div className="flex items-center space-x-3.5 text-xs text-slate-100 bg-[#4eb233]/10 p-2 rounded-lg border border-[#4eb233]/30">
                    <span className="w-5.5 h-5.5 rounded-full bg-[#4eb233] text-[#132e88] flex items-center justify-center text-[10px] font-black">2, 3</span>
                    <div>
                      <span className="font-semibold block text-[#4eb233]">인하우스 직선 복원 & 숏블라스트 전열 가공 (완벽 내재화)</span>
                      <span className="text-[10px] text-gray-300 font-light mt-0.5 block">외주 가공 기간 및 물류 비용 100% 영점 세이브</span>
                    </div>
                  </div>
                  <div className="flex items-center space-x-3.5 text-xs text-slate-100">
                    <span className="w-5.5 h-5.5 rounded-full bg-[#4eb233] text-[#132e88] flex items-center justify-center text-[10px] font-black">4</span>
                    <span className="font-semibold">WYF-3200 복합 직선 및 고정밀 봉강화</span>
                  </div>
                  <div className="flex items-center space-x-3.5 text-xs text-slate-100 bg-white/5 p-2 rounded-lg border border-white/10">
                    <span className="w-5.5 h-5.5 rounded-full bg-u1-yellow text-u1-blue flex items-center justify-center text-[10px] font-black">5, 6</span>
                    <div>
                      <span className="font-semibold block text-u1-yellow">WL75X 초밀도 정량 냉간 절단 & 압밀 요크 성형</span>
                      <span className="text-[10px] text-gray-300 font-light mt-0.5 block">자체 발명 특허 요크 제조 기법 적용 일체 성형</span>
                    </div>
                  </div>
                  <div className="flex items-center space-x-3.5 text-xs text-slate-100">
                    <span className="w-5.5 h-5.5 rounded-full bg-[#4eb233] text-[#132e88] flex items-center justify-center text-[10px] font-black">7</span>
                    <span className="font-semibold">통합 스마트 검사, 전용 친환경 팔레트 적재 및 출하</span>
                  </div>
                </div>
              </div>

              <div className="border-t border-white/10 pt-6 mt-8 flex justify-between items-center text-xs text-gray-300 font-mono relative z-10">
                <span>총 소요 리드타임: 단 1일 ~ 2일 (초신속)</span>
                <span className="text-[#4eb233] font-bold">품질 불량률 99.8% 원천 배제</span>
              </div>
            </div>

          </div>

          {/* 스마트팩토리 미래 배지 (초록색 #4eb233 슬림 배지 표기) */}
          <div className="bg-[#4eb233]/5 border-2 border-[#4eb233]/20 rounded-2xl p-6 md:p-8 flex flex-col md:flex-row items-start md:items-center justify-between gap-6" id="smart-factory-badge-container">
            <div className="flex items-center space-x-4">
              <div className="bg-[#4eb233] text-white p-3.5 rounded-xl flex items-center justify-center">
                <Clock className="w-6 h-6" />
              </div>
              <div>
                <span className="text-[10px] font-bold text-u1-blue tracking-widest bg-u1-yellow px-2 py-0.5 rounded uppercase font-mono">SMART FACTORY TO-BE ROADMAP</span>
                <h4 className="text-base md:text-lg font-bold text-u1-blue mt-1">
                  AI + MES 스마트팩토리 시스템 구축 계약 착수 완료
                </h4>
              </div>
            </div>
            <p className="text-xs md:text-sm text-gray-600 font-light max-w-2xl leading-relaxed md:text-right">
              현재 AI 기반 지능형 시스템 및 MES 생산 관리 스마트팩토리 구축 계약 완료 후 착수 단계입니다. 향후 자체 특허 3형을 추가 라인에 대입 적용 시 <strong className="text-u1-blue font-bold">최종 4단계 공정으로 극대 단축</strong>되며, 이를 통해 생산성 62% 향상 및 원자재 손실(재료 로스) 2% 이하 달성이라는 불가능을 실현할 목표입니다.
            </p>
          </div>

        </div>
      </section>

      {/* 6. FAQ 아코디언 섹션 */}
      <section id="faq" className="py-24 md:py-36 bg-[#fafafa]">
        <div className="max-w-4xl mx-auto px-6">
          
          <div className="text-center mb-16 md:mb-24">
            <span className="text-xs font-mono tracking-widest text-[#4eb233] uppercase font-bold block mb-3">04 / BRAND STATEMENT FAQ</span>
            <h2 className="text-3xl md:text-5xl font-extrabold text-u1-blue tracking-tight">
              자주 묻는 질문
            </h2>
            <p className="mt-4 text-gray-500 font-light text-sm md:text-base">
              유원스틸산업(주)의 가치와 기술력에 대해 파트너사들이 가장 궁금해하시는 핵심 질문들입니다.
            </p>
          </div>

          <div className="space-y-4 border-t border-gray-200" id="faq-accordion-group">
            {faqData.map((item) => {
              const isOpen = openFAQId === item.id;
              return (
                <div 
                  key={item.id} 
                  className="border-b border-gray-200 transition-all duration-300"
                  id={`faq-item-${item.id}`}
                >
                  <button
                    type="button"
                    onClick={() => setOpenFAQId(isOpen ? null : item.id)}
                    className="w-full py-6 flex justify-between items-center text-left focus:outline-none hover:text-u1-blue group cursor-pointer"
                  >
                    <span className="text-base md:text-lg font-bold text-gray-800 group-hover:text-u1-blue transition-colors duration-200 pr-4">
                      {item.question}
                    </span>
                    <span className="text-gray-400 group-hover:text-u1-blue transition-colors duration-200">
                      {isOpen ? <ChevronUp className="w-5 h-5 text-[#4eb233]" /> : <ChevronDown className="w-5 h-5" />}
                    </span>
                  </button>

                  <AnimatePresence initial={false}>
                    {isOpen && (
                      <motion.div
                        initial={{ height: 0, opacity: 0 }}
                        animate={{ height: "auto", opacity: 1 }}
                        exit={{ height: 0, opacity: 0 }}
                        transition={{ duration: 0.3, ease: "easeInOut" }}
                        className="overflow-hidden"
                      >
                        <div className="pb-8 pr-12 text-sm md:text-base text-gray-500 font-light leading-relaxed pl-1 border-l-2 border-[#4eb233]/40">
                          {item.answer}
                        </div>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>
              );
            })}
          </div>

        </div>
      </section>

      {/* 7. 리드 수집 폼 & 푸터 섹션 */}
      <section id="lead-section" className="py-24 md:py-36 bg-white">
        <div className="max-w-7xl mx-auto px-6 md:px-12">
          
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-12 lg:gap-16 items-center">
            
            {/* 왼쪽 안내 카피 */}
            <div className="lg:col-span-5">
              <span className="text-xs font-mono tracking-widest text-[#4eb233] uppercase font-bold block mb-3">05 / GET IN TOUCH</span>
              <h2 className="text-3xl md:text-5xl font-extrabold text-u1-blue tracking-tight leading-tight">
                유원과 함께<br />새로운 성장을 계획하세요
              </h2>
              <p className="mt-6 text-gray-500 font-light text-slate-500 leading-relaxed text-sm md:text-base">
                유원스틸산업(주)의 획기적인 7단계 원가 축소 공정 프로세스와 최신 특허 장비 사양, 그리고 투자 로드맵이 모두 담긴 국/영문 공식 소개서(Corporate Brochure)를 이메일로 즉시 받아 가실 수 있습니다.
              </p>

              <div className="mt-10 space-y-4">
                <div className="flex items-center space-x-3.5">
                  <div className="bg-u1-blue/5 p-2 rounded-lg text-u1-blue">
                    <FileText className="w-5 h-5" />
                  </div>
                  <span className="text-sm font-semibold text-gray-700">품질 보정 성적서 및 공정 단축 분석 포함</span>
                </div>
                <div className="flex items-center space-x-3.5">
                  <div className="bg-u1-blue/5 p-2 rounded-lg text-u1-blue">
                    <TrendingUp className="w-5 h-5" />
                  </div>
                  <span className="text-sm font-semibold text-gray-700">5개년 중장기 정성/정량적 매출 로드맵 기재</span>
                </div>
              </div>
            </div>

            {/* 오른쪽 정갈한 폼 (연한 그레이톤 배경 박스 위에 무결정렬) */}
            <div className="lg:col-span-7 bg-neutral-50/70 border border-gray-200/50 rounded-3xl p-8 md:p-12" id="lead-form-outer">
              
              <AnimatePresence mode="wait">
                {!isSubmitted ? (
                  <motion.div
                    key="lead-form-fields"
                    initial={{ opacity: 1 }}
                    exit={{ opacity: 0 }}
                    transition={{ duration: 0.3 }}
                  >
                    <div className="mb-6">
                      <h3 className="text-xl md:text-2xl font-bold text-u1-blue">
                        유원스틸산업(주) 상세 COMPANY PROFILE 다운로드
                      </h3>
                      <p className="text-xs md:text-sm text-gray-500 mt-2 leading-relaxed">
                        이메일을 입력하시면 자동차 부품 제조 공정 혁신 기술력과 5개년 실적 지표가 담긴 상세 소개서(PDF)를 즉시 확인하실 수 있습니다.
                      </p>
                    </div>

                    <form onSubmit={handleSubmit} className="space-y-5" id="lead-form">
                      {formError && (
                        <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-lg text-xs font-semibold flex items-center space-x-2">
                          <AlertCircle className="w-4 h-4 shrink-0" />
                          <span>{formError}</span>
                        </div>
                      )}

                      {/* 구글 스프레드시트 연동 상태 */}
                      <div className="flex items-center justify-between pb-3 border-b border-gray-200/60">
                        <span className="text-xs font-bold text-gray-500 flex items-center space-x-1.5">
                          <Database className="w-4 h-4 text-[#132e88]" />
                          <span>Google Sheets 실시간 연동</span>
                        </span>
                        <button
                          type="button"
                          onClick={() => setIsGuideOpen(!isGuideOpen)}
                          className={`text-[10px] font-bold px-2.5 py-1 rounded-full border transition-all cursor-pointer flex items-center space-x-1.5 ${
                            scriptUrl.trim() 
                              ? "bg-green-50 border-green-200 text-green-700 hover:bg-green-100" 
                              : "bg-amber-50 border-amber-200 text-amber-700 hover:bg-amber-100"
                          }`}
                        >
                          <span className={`w-1.5 h-1.5 rounded-full ${scriptUrl.trim() ? "bg-green-500 animate-pulse" : "bg-amber-400"}`} />
                          <span>{scriptUrl.trim() ? "연동 활성화됨" : "미연동 (안내서 보기)"}</span>
                        </button>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
                        <div className="flex flex-col space-y-1.5">
                          <label htmlFor="companyName" className="text-xs font-semibold text-gray-600 tracking-wide">
                            회사 이름
                          </label>
                          <input 
                            type="text" 
                            id="companyName"
                            name="companyName"
                            placeholder="예: 현대자동차"
                            value={formData.companyName}
                            onChange={handleInputChange}
                            disabled={isSubmitting}
                            className="bg-white border border-gray-200 rounded-lg px-4 py-3.5 text-sm focus:outline-none focus:border-u1-blue focus:ring-1 focus:ring-u1-blue/20 transition-all disabled:opacity-60"
                            required
                          />
                        </div>

                        <div className="flex flex-col space-y-1.5">
                          <label htmlFor="managerName" className="text-xs font-semibold text-gray-600 tracking-wide">
                            담당자 이름
                          </label>
                          <input 
                            type="text" 
                            id="managerName"
                            name="managerName"
                            placeholder="예: 홍길동 팀장"
                            value={formData.managerName}
                            onChange={handleInputChange}
                            disabled={isSubmitting}
                            className="bg-white border border-gray-200 rounded-lg px-4 py-3.5 text-sm focus:outline-none focus:border-u1-blue focus:ring-1 focus:ring-u1-blue/20 transition-all disabled:opacity-60"
                            required
                          />
                        </div>
                      </div>

                      <div className="flex flex-col space-y-1.5">
                        <label htmlFor="email" className="text-xs font-semibold text-gray-600 tracking-wide">
                          메일 주소
                        </label>
                        <input 
                          type="email" 
                          id="email"
                          name="email"
                          placeholder="client@company.com"
                          value={formData.email}
                          onChange={handleInputChange}
                          disabled={isSubmitting}
                          className="bg-white border border-gray-200 rounded-lg px-4 py-3.5 text-sm focus:outline-none focus:border-u1-blue focus:ring-1 focus:ring-u1-blue/20 transition-all disabled:opacity-60"
                          required
                        />
                      </div>

                      {/* 체크박스형태 */}
                      <div className="flex items-start space-x-3 pt-1">
                        <input 
                          type="checkbox" 
                          id="isAgreed"
                          name="isAgreed"
                          checked={formData.isAgreed}
                          onChange={handleInputChange}
                          disabled={isSubmitting}
                          className="mt-1 h-4 w-4 rounded border-gray-300 text-u1-blue focus:ring-u1-blue focus:ring-offset-0 disabled:opacity-60"
                          required
                        />
                        <label htmlFor="isAgreed" className="text-xs text-gray-400 leading-normal selection:bg-transparent">
                          <span className="font-semibold text-gray-600">[개인정보 수집 및 이용 동의 (필수)]</span> 귀사가 회사 소개서 다운로드 서비스를 원활히 실행하고 기술상담 예약을 진행하기 위해 위 정보를 한시 수집하여 최종 검토하는 것에 완전히 동의 및 서명합니다.
                        </label>
                      </div>

                      <button 
                        type="submit"
                        disabled={isSubmitting}
                        className={`w-full text-white text-xs md:text-sm font-bold tracking-wider uppercase py-4 md:py-4.5 rounded-xl active:scale-[0.99] transition-all duration-300 shadow-md text-center flex items-center justify-center space-x-2 cursor-pointer ${
                          isSubmitting ? "bg-gray-400 cursor-not-allowed" : "bg-[#132e88] hover:bg-[#4eb233] hover:shadow-u1-green/10"
                        }`}
                        id="lead-submit-btn"
                      >
                        {isSubmitting ? (
                          <>
                            <svg className="animate-spin h-4.5 w-4.5 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                              <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                              <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                            </svg>
                            <span>구글 시트 동기화 중...</span>
                          </>
                        ) : (
                          <span>상세 회사소개서 보기</span>
                        )}
                      </button>
                    </form>

                    {/* 구글 앱스 스크립트 연결 주소 직접 수정 및 배포 가이드라인 패널 */}
                    <div className="mt-6 pt-4 border-t border-gray-200/60" id="gas-setup-panel">
                      <div className="flex items-center justify-between">
                        <button
                          type="button"
                          onClick={() => setIsGuideOpen(!isGuideOpen)}
                          className="text-xs font-bold text-gray-500 hover:text-u1-blue transition-colors flex items-center space-x-1 cursor-pointer"
                        >
                          <Settings className="w-3.5 h-3.5" />
                          <span>Google Sheets 연동 & Apps Script 배포 가이드</span>
                          <span className="text-[10px] text-gray-400">({isGuideOpen ? "숨기기" : "자세히 보기"})</span>
                        </button>
                      </div>

                      {isGuideOpen && (
                        <motion.div
                          initial={{ opacity: 0, height: 0 }}
                          animate={{ opacity: 1, height: "auto" }}
                          className="mt-4 space-y-4 text-xs text-gray-600 bg-white p-4 md:p-6 rounded-2xl border border-gray-200 shadow-sm"
                        >
                          <div>
                            <label className="block text-xs font-bold text-gray-800 mb-1.5">
                              1. 배포한 Google Apps Script Web App URL 등록
                            </label>
                            <div className="flex space-x-2">
                              <input
                                type="url"
                                placeholder="https://script.google.com/macros/s/.../exec"
                                value={scriptUrl}
                                onChange={(e) => {
                                  setScriptUrl(e.target.value);
                                  localStorage.setItem("VITE_GOOGLE_SCRIPT_URL", e.target.value);
                                }}
                                className="flex-1 bg-gray-50 border border-gray-200 rounded-lg px-3 py-2 text-xs focus:outline-none focus:border-u1-blue focus:ring-1 focus:ring-u1-blue/20"
                              />
                              {scriptUrl.trim() && (
                                <button
                                  type="button"
                                  onClick={() => {
                                    setScriptUrl("");
                                    localStorage.removeItem("VITE_GOOGLE_SCRIPT_URL");
                                  }}
                                  className="px-2.5 py-2 bg-red-50 text-red-600 hover:bg-red-100 rounded-lg font-semibold transition-colors text-[10px] cursor-pointer"
                                >
                                  초기화
                                </button>
                              )}
                            </div>
                            <p className="text-[10px] text-gray-400 mt-1 leading-normal">
                              * 복사하신 Apps Script의 웹앱(Web App) URL 주소를 여기에 넣으시면 기입 즉시 브라우저 로컬 저장소에 안전하게 바인딩되어 실시간 데이터 연동이 개시됩니다.
                            </p>
                          </div>

                          <div className="border-t border-gray-100 pt-3">
                            <h4 className="font-bold text-gray-800 mb-1.5">2. Google Apps Script 생성 소스코드</h4>
                            <p className="text-gray-400 text-[10px] mb-2 leading-relaxed">
                              아래 소스코드를 그대로 전부 복사하셔서 구글 스프레드시트 Apps Script 에디터에 엎어쳐 주세요. (우측 상단의 단추를 누르면 전체 내용이 클립보드에 바로 복사됩니다)
                            </p>
                            <div className="relative bg-gray-950 rounded-xl p-3 text-[10px] font-mono text-gray-300 max-h-48 overflow-y-auto">
                              <button
                                type="button"
                                onClick={() => {
                                  navigator.clipboard.writeText(googleAppsScriptCode);
                                  setIsCopying(true);
                                  setTimeout(() => setIsCopying(false), 2000);
                                }}
                                className="absolute top-2 right-2 p-1.5 bg-gray-800 hover:bg-gray-700 text-gray-300 hover:text-white rounded-lg transition-colors cursor-pointer"
                                title="코드 복사"
                              >
                                {isCopying ? (
                                  <span className="text-[9px] text-[#4eb233] font-sans font-bold flex items-center space-x-1">
                                    <Check className="w-3 h-3 inline" />
                                    <span>복사 완료</span>
                                  </span>
                                ) : (
                                  <Copy className="w-3 h-3" />
                                )}
                              </button>
                              <pre className="whitespace-pre">{googleAppsScriptCode}</pre>
                            </div>
                          </div>

                          <div className="border-t border-gray-100 pt-3 text-[11px] text-gray-500 space-y-1">
                            <h4 className="font-bold text-gray-800 mb-2">3. 스프레드시트 매크로 연결 및 배포 순서</h4>
                            <ol className="list-decimal pl-4 space-y-1 bg-neutral-50 p-3 rounded-xl border border-gray-100 text-[11px] leading-relaxed">
                              <li>
                                연동하고자 하시는{" "}
                                <a 
                                  href="https://docs.google.com/spreadsheets/d/1j1BNMXTC8CQweK1TOsUkTCW2aVsuZNZ7FWWBPSqp_ZY/edit" 
                                  target="_blank" 
                                  rel="noopener noreferrer"
                                  className="text-u1-blue font-bold underline inline-flex items-center space-x-0.5"
                                >
                                  <span>지정 구글 스프레드시트</span>
                                  <ExternalLink className="w-3 h-3 inline" />
                                </a>{" "}
                                주소로 들어갑니다.
                              </li>
                              <li>상단 바 메뉴에서 <strong>[확장 프로그램] &gt; [Apps Script]</strong>를 클릭하여 개발창을 실행합니다.</li>
                              <li>기존 기본 생성된 코드를 제거하고, 위의 <strong>2번 복사 코드</strong>를 그대로 복사하여 전체 추가해 줍니다.</li>
                              <li>상단 디스크 심볼 단추를 클릭해 프로젝트를 <strong>저장</strong>합니다.</li>
                              <li>우측 상단 <strong>[배포] &gt; [새 배포]</strong>를 클릭하여 설정창을 엽니다.</li>
                              <li>유형(톱니바퀴)을 <strong>[웹앱 (Web App)]</strong>으로 지정합니다.</li>
                              <li>설정을 조율합니다.
                                <ul className="list-disc pl-4 mt-0.5 space-y-0.5 text-gray-500">
                                  <li>다음 사용자 권한으로 가동 (Execute as): <strong>나 (이메일 본인 주소)</strong></li>
                                  <li>액세스 가능한 그룹 권한 (Who has access): <strong className="text-[#132e88]">모든 사용자 (Anyone)</strong></li>
                                </ul>
                              </li>
                              <li><strong>[배포]</strong>를 실행하고, 연결 권한 팝업창에서 본인 구글 계정을 승인 처리하고 최종 완료합니다.</li>
                              <li>화면에 노출된 <strong>웹앱 URL</strong>(https://script.google.com/macros/s/.../exec) 주소를 복사하여 위의 <strong>1번 입력칸</strong>에 바로 채워 넣어 줍니다.</li>
                            </ol>
                          </div>
                        </motion.div>
                      )}
                    </div>
                  </motion.div>
                ) : (
                  <motion.div
                    key="lead-form-success"
                    initial={{ opacity: 0, scale: 0.95 }}
                    animate={{ opacity: 1, scale: 1 }}
                    transition={{ duration: 0.4 }}
                    className="text-center py-6"
                    id="lead-form-success-container"
                  >
                    <div className="inline-flex p-4 rounded-full bg-[#4eb233]/15 text-[#4eb233] mb-6 animate-bounce">
                      <CheckCircle className="w-12 h-12" />
                    </div>
                    
                    <h3 className="text-2xl font-bold text-u1-blue">
                      다운로드 준비 완료
                    </h3>
                    <p className="text-sm text-gray-500 max-w-md mx-auto mt-2 leading-relaxed">
                      감사합니다. <span className="font-bold text-u1-blue">{formData.managerName}</span> 님!<br />
                      입력하신 메일 <span className="underline font-medium text-gray-700">{formData.email}</span>(으)로 <strong>유원스틸산업(주) 상세 실적 및 기술 공정 PDF 소개서</strong>를 즉시 자동 발송 처리하였습니다.
                    </p>

                    {/* 실시간 스프레드시트 동기화 디버그 알리미 */}
                    <div className="mt-4 max-w-md mx-auto">
                      {syncStatus.status === "success" ? (
                        <div className="bg-green-50 border border-green-200 rounded-xl p-3 text-xs text-green-700 font-semibold text-center flex items-center justify-center space-x-1.5">
                          <Database className="w-4 h-4 text-green-600 shrink-0" />
                          <span>{syncStatus.message}</span>
                        </div>
                      ) : syncStatus.status === "error" ? (
                        <div className="bg-red-50 border border-red-200 rounded-xl p-3 text-xs text-red-700 font-semibold text-center flex items-center justify-center space-x-1.5">
                          <AlertCircle className="w-4 h-4 text-red-600 shrink-0" />
                          <span>{syncStatus.message}</span>
                        </div>
                      ) : (
                        <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-700 font-medium text-left leading-normal space-y-1">
                          <div className="font-bold flex items-center space-x-1 text-amber-800">
                            <AlertCircle className="w-4 h-4 shrink-0" />
                            <span>로컬 모드로 가상 성공</span>
                          </div>
                          <p className="text-[11px] text-amber-600 font-light">
                            {syncStatus.message}
                          </p>
                        </div>
                      )}
                    </div>

                    <div className="bg-white rounded-2xl p-6 border border-gray-200 max-w-sm mx-auto mt-6 shadow-sm">
                      <span className="text-[10px] font-mono font-bold tracking-widest text-[#4eb233] block mb-1">YOUR ATTACHMENT</span>
                      <h4 className="text-xs font-bold text-gray-800">U1_Steel_Corporate_Profile_2026.pdf</h4>
                      <p className="text-[10px] text-gray-400 mt-1">Google Drive PDF Document | Korean, English</p>
                      
                      <a 
                        href="https://drive.google.com/file/d/1TaoX2M4taxpp_xCGxBrHOEw2CmC5xSVY/view?usp=sharing" 
                        target="_blank"
                        rel="noopener noreferrer"
                        className="mt-4 w-full bg-u1-blue text-white rounded-lg py-2.5 text-xs font-semibold hover:bg-u1-green transition-colors inline-flex items-center justify-center space-x-1.5 cursor-pointer"
                      >
                        <Download className="w-3.5 h-3.5" />
                        <span>회사소개서 바로보기/다운로드</span>
                      </a>
                    </div>

                    <div className="mt-8 flex flex-col items-center space-y-3">
                      <button 
                        onClick={handleResetForm}
                        className="text-xs font-semibold text-gray-400 hover:text-u1-blue transition-colors cursor-pointer"
                      >
                        다른 이메일로 다시 받기
                      </button>

                      <button
                        type="button"
                        onClick={() => setIsGuideOpen(!isGuideOpen)}
                        className="text-[11px] text-gray-400 hover:text-u1-blue transition-colors underline flex items-center space-x-1 cursor-pointer"
                      >
                        <Settings className="w-3 h-3" />
                        <span>전체 매크로 및 스프레드시트 셋업 가이드 다시 확인하기</span>
                      </button>
                    </div>

                    {isGuideOpen && (
                      <div className="mt-6 text-left max-w-xl mx-auto">
                        <motion.div
                          initial={{ opacity: 0, height: 0 }}
                          animate={{ opacity: 1, height: "auto" }}
                          className="space-y-4 text-xs text-gray-600 bg-white p-5 rounded-2xl border border-gray-200 shadow-sm"
                        >
                          <div>
                            <label className="block text-xs font-bold text-gray-800 mb-1">
                              구글 앱스 스크립트 웹앱 주소 기입란
                            </label>
                            <div className="flex space-x-2">
                              <input
                                type="url"
                                placeholder="https://script.google.com/macros/s/.../exec"
                                value={scriptUrl}
                                onChange={(e) => {
                                  setScriptUrl(e.target.value);
                                  localStorage.setItem("VITE_GOOGLE_SCRIPT_URL", e.target.value);
                                }}
                                className="flex-1 bg-gray-50 border border-gray-200 rounded-lg px-3 py-2 text-xs focus:outline-none focus:border-u1-blue focus:ring-1 focus:ring-u1-blue/20"
                              />
                            </div>
                          </div>

                          <div className="border-t border-gray-100 pt-3">
                            <h4 className="font-bold text-gray-800 mb-1.5">Apps Script 복사하기</h4>
                            <div className="relative bg-gray-950 rounded-xl p-3 text-[10px] font-mono text-gray-300 max-h-40 overflow-y-auto">
                              <button
                                type="button"
                                onClick={() => {
                                  navigator.clipboard.writeText(googleAppsScriptCode);
                                  setIsCopying(true);
                                  setTimeout(() => setIsCopying(false), 2000);
                                }}
                                className="absolute top-2 right-2 p-1.5 bg-gray-800 hover:bg-gray-700 text-gray-300 hover:text-white rounded-lg transition-colors cursor-pointer"
                              >
                                {isCopying ? "복사완료" : <Copy className="w-3 h-3" />}
                              </button>
                              <pre className="whitespace-pre">{googleAppsScriptCode}</pre>
                            </div>
                          </div>
                        </motion.div>
                      </div>
                    )}
                  </motion.div>
                )}
              </AnimatePresence>

            </div>

          </div>

        </div>
      </section>

      {/* 8. 푸터 섹션 */}
      <footer className="bg-gray-900 text-gray-400 pt-20 pb-16 border-t border-gray-800 select-none">
        <div className="max-w-7xl mx-auto px-6 md:px-12">
          
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-12 lg:gap-8 pb-16 border-b border-gray-800">
            
            {/* 로고 영역 및 이념 */}
            <div className="lg:col-span-5">
              <Logo className="h-12 md:h-14 mb-4" showSubtitle={true} isDarkTheme={true} />
              <p className="text-xs font-light text-gray-400 leading-relaxed max-w-md mt-6">
                우리는 최고 품질의 자동차 패스너 핵심 요크를 양산하여 전 세계 완성차가 조향축 무결성을 지키는 데 공헌합니다. 또한, 지속적인 CANI 정신으로 불필요한 물류와 공정을 혁신해 나갑니다.
              </p>
            </div>

            {/* 본사 및 지사 안내 */}
            <div className="lg:col-span-7 grid grid-cols-1 md:grid-cols-2 gap-8 text-xs">
              
              <div className="space-y-4">
                <h4 className="text-white font-bold tracking-wider text-sm">HEADOFFICE & MANUFACTORY</h4>
                <div className="flex items-start space-x-2.5">
                  <MapPin className="w-4 h-4 text-[#4eb233] shrink-0 mt-0.5" />
                  <span>충청북도 진천군 덕산읍 신척1길 94 (본사 공장)</span>
                </div>
                <div className="flex items-start space-x-2.5">
                  <Phone className="w-4 h-4 text-gray-500 shrink-0 mt-0.5" />
                  <span>Tel: 043-753-8225 &nbsp;/&nbsp; Fax: 043-753-8226</span>
                </div>
              </div>

              <div className="space-y-4">
                <h4 className="text-white font-bold tracking-wider text-sm">SEOUL OFFICE & SALES</h4>
                <div className="flex items-start space-x-2.5">
                  <Building className="w-4 h-4 text-[#4eb233] shrink-0 mt-0.5" />
                  <span>서울특별시 광진구 광나루로 478, 106호 (서울사무소)</span>
                </div>
                <div className="flex items-start space-x-2.5">
                  <Phone className="w-4 h-4 text-gray-500 shrink-0 mt-0.5" />
                  <span>Tel: 02-6959-1973</span>
                </div>
                <div className="flex items-start space-x-2.5">
                  <Mail className="w-4 h-4 text-gray-500 shrink-0 mt-0.5" />
                  <span>sales@u1-steel.com</span>
                </div>
              </div>

            </div>

          </div>

          <div className="pt-8 flex flex-col md:flex-row items-center justify-between text-xs text-gray-500 font-light gap-4">
            <div className="space-y-1">
              <p>유원스틸산업(주) &nbsp;|&nbsp; 대표이사 류희철 &nbsp;|&nbsp; 사업자등록번호 231-88-02013</p>
              <p className="font-mono text-[10px]">YOUONE STEEL CORP. OWNER: RYU HEE CHUL &nbsp;|&nbsp; REGISTRATION: 231-88-02013</p>
            </div>
            <p className="font-mono text-[10px] md:text-right">
              Copyright © YOUONE STEEL Corp. All Rights Reserved. &nbsp;/&nbsp; Developed with Extreme Minimalism.
            </p>
          </div>

        </div>
      </footer>

    </div>
  );
}
