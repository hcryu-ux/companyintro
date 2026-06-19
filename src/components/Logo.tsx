import { SVGProps } from "react";

export default function Logo({ 
  className = "h-9", 
  showSubtitle = true, 
  isDarkTheme = false, 
  ...props 
}: SVGProps<SVGSVGElement> & { 
  className?: string; 
  showSubtitle?: boolean; 
  isDarkTheme?: boolean; 
}) {
  const textColor = isDarkTheme ? "#ffffff" : "#3e3a39";
  const subColor = isDarkTheme ? "#cccccc" : "#3e3a39";

  return (
    <div className={`flex items-center select-none ${className}`}>
      {/* 3D-like Minimalistic Premium SVG Symbol representing U1 Steel Wave */}
      <svg
        viewBox="0 15 380 170"
        className="h-full w-auto object-contain"
        xmlns="http://www.w3.org/2000/svg"
        {...props}
      >
        <g id="U1-Symbol" transform="translate(10, 10)">
          {/* Blue Wave (Leftmost) */}
          <path
            d="M5,115 C20,115 28,95 44,118 C50,126 50,135 40,140 C30,145 15,142 5,115 Z"
            fill="url(#blueGrad)"
          />
          <circle cx="21" cy="95" r="9" fill="#009fe3" />

          {/* Green Wave (Middle) */}
          <path
            d="M36,114 C50,105 75,100 115,102 C108,135 90,160 70,165 C55,168 45,140 36,114 Z"
            fill="#4eb233"
          />
          <circle cx="70" cy="72" r="12" fill="#4eb233" />

          {/* Yellow Wave (Right/Main) */}
          <path
            d="M102,74 C100,105 85,145 92,175 C102,175 140,130 168,54 C130,64 112,68 102,74 Z"
            fill="#fdd000"
          />
          <circle cx="134" cy="38" r="15" fill="#fdd000" />
        </g>

        {/* Brand Text Elements (Rendered via SVG text for high-fidelity vector alignment) */}
        <g id="U1-Brand-Text">
          {/* "U1" - Heavy bold charcoal */}
          <text
            x="184"
            y="98"
            fontFamily="Pretendard, -apple-system, sans-serif"
            fontWeight="900"
            fontSize="54"
            fill={textColor}
            letterSpacing="-3"
          >
            U1
          </text>
          {/* "Steel" - Semi-bold charcoal */}
          <text
            x="248"
            y="98"
            fontFamily="Pretendard, -apple-system, sans-serif"
            fontWeight="700"
            fontSize="54"
            fill={textColor}
            letterSpacing="-1.5"
          >
            Steel
          </text>

          {/* Subtitle / Brand Philosophy - "당신이 1등이 되는 가치 유원" */}
          {showSubtitle && (
            <text
              x="184"
              y="138"
              fontFamily="Pretendard, -apple-system, sans-serif"
              fontWeight="700"
              fontSize="16.5"
              fill={subColor}
              letterSpacing="-0.5"
            >
              당신이 1등이 되는 가치 유원
            </text>
          )}
        </g>

        {/* Gradients definition */}
        <defs>
          <linearGradient id="blueGrad" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stopColor="#009fe3" />
            <stop offset="100%" stopColor="#132e88" />
          </linearGradient>
        </defs>
      </svg>
    </div>
  );
}
