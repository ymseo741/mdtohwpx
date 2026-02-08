import React, { useState, useCallback } from 'react';
import axios from 'axios';
import { motion, AnimatePresence } from 'framer-motion';
import {
    FileText,
    Upload,
    Download,
    RefreshCw,
    CheckCircle2,
    AlertCircle,
    FileDown,
    ChevronRight,
    Code
} from 'lucide-react';
import { useDropzone } from 'react-dropzone';
import ReactMarkdown from 'react-markdown';
import { clsx } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs) {
    return twMerge(clsx(inputs));
}

const App = () => {
    const [markdown, setMarkdown] = useState('');
    const [isConverting, setIsConverting] = useState(false);
    const [status, setStatus] = useState('idle');
    const [error, setError] = useState('');

    const onDrop = useCallback((acceptedFiles) => {
        const file = acceptedFiles[0];
        const reader = new FileReader();
        reader.onload = () => {
            setMarkdown(reader.result);
        };
        reader.readAsText(file);
    }, []);

    const { getRootProps, getInputProps, isDragActive } = useDropzone({
        onDrop,
        accept: { 'text/markdown': ['.md'] },
        multiple: false
    });

    const handleConvert = async () => {
        if (!markdown) return;
        setIsConverting(true);
        setStatus('idle');
        setError('');

        try {
            const formData = new FormData();
            formData.append('text', markdown);

            const response = await axios.post('/api/convert', formData, {
                responseType: 'blob'
            });

            const url = window.URL.createObjectURL(new Blob([response.data]));
            const link = document.createElement('a');
            link.href = url;
            link.setAttribute('download', 'proposal.hwpx');
            document.body.appendChild(link);
            link.click();
            link.remove();

            setStatus('success');
        } catch (err) {
            console.error(err);
            setStatus('error');
            setError(err.response?.data?.detail || 'An error occurred during conversion');
        } finally {
            setIsConverting(false);
        }
    };

    const loadRDTemplates = () => {
        setMarkdown(`# 연구개발계획서

## 과제 개요 및 개발 목표
**과제명**: 지능형 해상 충돌 방지 시스템(I-SCAS) 고도화  
**신청 기업**: (주)안티그래비티 | **기간**: 2026.03 ~ 2027.02

**핵심 개발 목표**:  
- AI 기반 충돌 위험 객체 식별 정확도 99% 달성
- 다중 선박 실시간 경로 예측 오차 범위 5m 이내 확보

## 개발 방법 및 추진 전략
**핵심 기술 구성**:  
\`\`\`
- 기술1: Transformer 기반 시계열 선박 궤적 예측 알고리즘
- 기술2: 고해상도 LiDAR 및 열화상 카메라 센서 퓨전
\`\`\`

## 성과 목표 및 사업화 로드맵
| 지표 | 목표치 | 검증방법 | 달성시점 |
|------|--------|----------|----------|
| 객체 식별률 | 99% | 자체 데이터셋 테스트 | 2026.12 |
`);
    };

    return (
        <div className="min-h-screen relative overflow-hidden flex flex-col items-center py-12 px-4">
            {/* Animated Background blobs */}
            <div className="animated-bg">
                <motion.div
                    animate={{ x: [0, 100, 0], y: [0, 50, 0] }}
                    transition={{ duration: 20, repeat: Infinity, ease: "linear" }}
                    className="blob top-[-10%] left-[-10%]"
                />
                <motion.div
                    animate={{ x: [0, -80, 0], y: [0, 100, 0] }}
                    transition={{ duration: 25, repeat: Infinity, ease: "linear" }}
                    className="blob bottom-[-10%] right-[-10%] bg-purple-500/10"
                />
            </div>

            {/* Header */}
            <motion.div
                initial={{ opacity: 0, y: -20 }}
                animate={{ opacity: 1, y: 0 }}
                className="text-center mb-12 relative z-10"
            >
                <div className="inline-flex items-center space-x-2 px-3 py-1 rounded-full bg-blue-500/10 border border-blue-500/20 text-blue-400 text-sm font-medium mb-4">
                    <span className="relative flex h-2 w-2">
                        <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-blue-400 opacity-75"></span>
                        <span className="relative inline-flex rounded-full h-2 w-2 bg-blue-500"></span>
                    </span>
                    <span>RD Business Proposal Tool</span>
                </div>
                <h1 className="text-5xl md:text-6xl font-['Outfit'] font-bold text-white mb-4 tracking-tight">
                    Markdown <span className="text-gradient">to HWPX</span>
                </h1>
                <p className="text-slate-400 text-lg max-w-2xl mx-auto">
                    마크다운 기반의 사업계획서를 고품질 HWPX 문서로 즉시 변환하세요.
                    모든 스타일링이 자동화되어 문서 작성 시간을 단축합니다.
                </p>
            </motion.div>

            {/* Main Container */}
            <div className="w-full max-w-6xl grid grid-cols-1 lg:grid-cols-2 gap-8 relative z-10">

                {/* Editor Side */}
                <motion.div
                    initial={{ opacity: 0, x: -20 }}
                    animate={{ opacity: 1, x: 0 }}
                    transition={{ delay: 0.1 }}
                    className="flex flex-col space-y-4"
                >
                    <div className="glass-card rounded-2xl p-6 flex-1 flex flex-col min-h-[500px]">
                        <div className="flex items-center justify-between mb-4">
                            <div className="flex items-center space-x-2">
                                <Code size={20} className="text-blue-400" />
                                <span className="text-white font-medium">Markdown Editor</span>
                            </div>
                            <button
                                onClick={loadRDTemplates}
                                className="text-xs text-blue-400 hover:text-blue-300 transition-colors flex items-center"
                            >
                                R&D 템플릿 불러오기 <ChevronRight size={14} />
                            </button>
                        </div>

                        <textarea
                            className="flex-1 w-full bg-slate-900/50 border border-slate-700/50 rounded-xl p-4 text-slate-300 font-mono text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/50 transition-all resize-none"
                            placeholder="# 여기에 내용을 입력하거나 파일을 드래그하세요..."
                            value={markdown}
                            onChange={(e) => setMarkdown(e.target.value)}
                        />

                        <div
                            {...getRootProps()}
                            className={cn(
                                "mt-4 border-2 border-dashed rounded-xl p-6 text-center transition-all cursor-pointer",
                                isDragActive ? "border-blue-500 bg-blue-500/10" : "border-slate-700 hover:border-slate-600 hover:bg-slate-800/30"
                            )}
                        >
                            <input {...getInputProps()} />
                            <Upload className="mx-auto mb-2 text-slate-500" size={24} />
                            <p className="text-sm text-slate-400">
                                {isDragActive ? "파일을 놓으세요" : "마크다운(MD) 파일을 드래그하거나 클릭하여 업로드"}
                            </p>
                        </div>
                    </div>
                </motion.div>

                {/* Preview & Action Side */}
                <motion.div
                    initial={{ opacity: 0, x: 20 }}
                    animate={{ opacity: 1, x: 0 }}
                    transition={{ delay: 0.2 }}
                    className="flex flex-col space-y-4"
                >
                    <div className="glass-card rounded-2xl p-6 flex-1 flex flex-col min-h-[500px] overflow-hidden">
                        <div className="flex items-center space-x-2 mb-4">
                            <FileText size={20} className="text-purple-400" />
                            <span className="text-white font-medium">Live Preview</span>
                        </div>

                        <div className="flex-1 bg-white rounded-xl p-6 text-slate-800 overflow-y-auto prose prose-sm max-w-none">
                            {markdown ? (
                                <ReactMarkdown>{markdown}</ReactMarkdown>
                            ) : (
                                <div className="h-full flex items-center justify-center text-slate-400 italic">
                                    미리보기가 여기에 표시됩니다.
                                </div>
                            )}
                        </div>

                        <button
                            onClick={handleConvert}
                            disabled={!markdown || isConverting}
                            className={cn(
                                "mt-6 w-full py-4 rounded-xl flex items-center justify-center space-x-2 font-bold transition-all shadow-lg",
                                !markdown || isConverting
                                    ? "bg-slate-800 text-slate-500 cursor-not-allowed"
                                    : "premium-gradient text-white hover:scale-[1.02] active:scale-[0.98] hover:shadow-blue-500/25"
                            )}
                        >
                            {isConverting ? (
                                <RefreshCw className="animate-spin" size={20} />
                            ) : (
                                <FileDown size={20} />
                            )}
                            <span>{isConverting ? "변환 중..." : "HWPX 문서로 내보내기"}</span>
                        </button>

                        <AnimatePresence>
                            {status === 'success' && (
                                <motion.div
                                    initial={{ opacity: 0, y: 10 }}
                                    animate={{ opacity: 1, y: 0 }}
                                    exit={{ opacity: 0 }}
                                    className="mt-4 p-3 bg-green-500/10 border border-green-500/20 rounded-lg flex items-center space-x-2 text-green-400 text-sm"
                                >
                                    <CheckCircle2 size={16} />
                                    <span>변환이 완료되었습니다! 다운로드가 시작됩니다.</span>
                                </motion.div>
                            )}
                            {status === 'error' && (
                                <motion.div
                                    initial={{ opacity: 0, y: 10 }}
                                    animate={{ opacity: 1, y: 0 }}
                                    exit={{ opacity: 0 }}
                                    className="mt-4 p-3 bg-red-500/10 border border-red-500/20 rounded-lg flex items-center space-x-2 text-red-400 text-sm"
                                >
                                    <AlertCircle size={16} />
                                    <span className="truncate">{error}</span>
                                </motion.div>
                            )}
                        </AnimatePresence>
                    </div>
                </motion.div>
            </div>

            {/* Footer info */}
            <motion.div
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                transition={{ delay: 0.5 }}
                className="mt-12 text-slate-500 text-sm italic"
            >
                © 2026 Antigravity - Powered by Next-gen Agentic AI
            </motion.div>
        </div>
    );
};

export default App;
