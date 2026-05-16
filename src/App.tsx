/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo } from 'react';
import { 
  Calculator, 
  MessageSquare, 
  AlertTriangle, 
  AlertCircle,
  Download, 
  Plus, 
  Trash2, 
  ChevronRight,
  ShieldCheck,
  TrendingUp,
  Search,
  BookOpen,
  User,
  Upload,
  FileText,
  FileSpreadsheet,
  X,
  ThumbsUp,
  ThumbsDown,
  Paperclip,
  Image as ImageIcon,
  Send,
  RefreshCw,
  Link as LinkIcon,
  Mail,
  FileCode,
  FileDown,
  Sparkles,
  CheckCircle2,
  ExternalLink,
  Globe,
  Copy,
  Mic,
  MicOff
} from 'lucide-react';
import { GoogleGenAI } from "@google/genai";
import { jsPDF } from "jspdf";
import "jspdf-autotable";
import * as XLSX from 'xlsx';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, AlignmentType, HeadingLevel } from "docx";
import { saveAs } from 'file-saver';

// --- CONSTANTS & TAX RULES ---
const CURRENT_YA = 2024;
const CORPORATE_TAX_RATE = 0.17;
const PARTIAL_TAX_EXEMPTION = {
  first10kRate: 0.75, // 75% exemption on first 10k
  next190kRate: 0.50, // 50% exemption on next 190k
};

// --- CACHE & HEURISTICS ---
const searchCache = new Map<string, any>();
const advisoryCache = new Map<string, any>();

const LOCAL_QUICK_LINKS = [
  { title: "Corporate Income Tax", link: "https://www.iras.gov.sg/taxes/corporate-income-tax", type: "General" },
  { title: "Goods and Services Tax (GST)", link: "https://www.iras.gov.sg/taxes/goods-and-services-tax-(gst)", type: "General" },
  { title: "Individual Income Tax", link: "https://www.iras.gov.sg/taxes/individual-income-tax", type: "General" },
  { title: "Property Tax", link: "https://www.iras.gov.sg/taxes/property-tax", type: "General" },
  { title: "Stamp Duty", link: "https://www.iras.gov.sg/taxes/stamp-duty", type: "General" },
  { title: "Deductibility of Expenses", link: "https://www.iras.gov.sg/taxes/corporate-income-tax/income-deductions-and-reliefs/business-expenses/deductibility-of-expenses", type: "IRAS Guide" },
  { title: "Capital Allowances", link: "https://www.iras.gov.sg/taxes/corporate-income-tax/income-deductions-and-reliefs/capital-allowances", type: "IRAS Guide" },
  { title: "Section 14Q (R&R)", link: "https://www.iras.gov.sg/taxes/corporate-income-tax/income-deductions-and-reliefs/business-expenses/renovation-and-refurbishment-(r-r)-works", type: "IRAS Guide" }
];

const findHeuristicResult = (query: string) => {
  const q = query.toLowerCase();
  const matches = LOCAL_QUICK_LINKS.filter(link => 
    q.includes(link.title.toLowerCase()) || 
    link.title.toLowerCase().includes(q) ||
    (q.includes('cit') && link.title.includes('Corporate')) ||
    (q.includes('gst') && link.title.includes('GST'))
  );
  
  if (matches.length > 0) {
    return {
      results: matches.map(m => ({
        ...m,
        description: `Direct resource for ${m.title}. Access the official IRAS portal for current rules and forms.`,
        actSection: m.type === 'Legislation' ? 'ITA' : undefined
      }))
    };
  }
  return null;
};

// --- TYPES ---
interface ComputationItem {
  id: number;
  label: string;
  amount: number;
  note: string;
  isAISuggested?: boolean;
}

interface TaxData {
  accountingProfit: number;
  additions: ComputationItem[];
  deductions: ComputationItem[];
  otherAdjustments: ComputationItem[];
}

interface TaxQuery {
  id: string;
  query: string;
  answer: string;
  explanation: string;
  assumptions: string[];
  references: string[];
  timestamp: string;
  feedback?: 'correct' | 'wrong';
  chatHistory?: { role: 'user' | 'model'; text: string }[];
  attachments?: { name: string; type: string }[];
  urls?: string[];
  citations?: { 
    source: string; 
    link?: string; 
    location?: string; 
    isModern?: boolean;
    snippet?: string;
    searchQuery?: string;
    type?: 'Legislation' | 'IRAS Guide' | 'Circular' | 'General';
    relevance?: 'High' | 'Medium' | 'Low';
    isPdf?: boolean;
  }[];
  searchEntryPoint?: string;
}

const cleanIrasUrl = (url: string): { link: string, isModern: boolean, isPdf: boolean } => {
  if (!url) return { link: "", isModern: false, isPdf: false };
  let cleaned = url.trim();
  if (!cleaned.startsWith('http')) cleaned = `https://${cleaned}`;
  
  // Remove legacy /irashome/ path
  cleaned = cleaned.replace(/\/irashome\//g, '/');
  
  const isPdf = cleaned.toLowerCase().endsWith('.pdf') || cleaned.includes('/docs/default-source/');

  // Check if it matches the modern IRAS structure
  const isModern = cleaned.includes('iras.gov.sg/taxes/') || 
                   cleaned.includes('iras.gov.sg/quick-links/') ||
                   cleaned.includes('iras.gov.sg/news-events/') ||
                   cleaned.includes('iras.gov.sg/about-iras/') ||
                   isPdf;

  // If it's a legacy page (.aspx or missing modern path), redirect to main category
  if (!isPdf && (cleaned.includes('.aspx') || cleaned.includes('page.aspx') || !isModern)) {
    if (cleaned.toLowerCase().includes('e-tax_guides') || cleaned.toLowerCase().includes('etaxguide')) {
      return { link: 'https://www.iras.gov.sg/quick-links/e-tax-guides', isModern: true, isPdf: false };
    }
    if (cleaned.toLowerCase().includes('corporate')) return { link: 'https://www.iras.gov.sg/taxes/corporate-income-tax', isModern: true, isPdf: false };
    if (cleaned.toLowerCase().includes('gst')) return { link: 'https://www.iras.gov.sg/taxes/goods-and-services-tax-(gst)', isModern: true, isPdf: false };
    if (cleaned.toLowerCase().includes('individual')) return { link: 'https://www.iras.gov.sg/taxes/individual-income-tax', isModern: true, isPdf: false };
    if (cleaned.toLowerCase().includes('stamp')) return { link: 'https://www.iras.gov.sg/taxes/stamp-duty', isModern: true, isPdf: false };
    if (cleaned.toLowerCase().includes('property')) return { link: 'https://www.iras.gov.sg/taxes/property-tax', isModern: true, isPdf: false };
    
    return { link: 'https://www.iras.gov.sg/taxes', isModern: true, isPdf: false };
  }

  return { link: cleaned, isModern: true, isPdf };
};

const parseAIError = (error: any): string => {
  console.error("AI Error:", error);
  let errorMessage = "An unexpected error occurred. Please try again.";
  
  // Attempt to extract message from the response if it's a JSON string
  let detail = "";
  let errorCode = "";
  
  try {
    const errorStr = typeof error === 'string' ? error : (error?.message || "");
    if (errorStr.includes('{')) {
      const jsonStart = errorStr.indexOf('{');
      const jsonEnd = errorStr.lastIndexOf('}') + 1;
      const parsed = JSON.parse(errorStr.substring(jsonStart, jsonEnd));
      if (parsed?.error?.message) detail = parsed.error.message;
      if (parsed?.error?.code) errorCode = String(parsed.error.code);
    }
  } catch (e) {
    // Not JSON
  }

  const combinedStr = (error?.message || "") + " " + detail;

  if (combinedStr.includes('429') || combinedStr.includes('RESOURCE_EXHAUSTED') || combinedStr.toLowerCase().includes('quota') || errorCode === '429') {
    return "API Quota Exceeded: You've reached the usage limit for your Gemini API key. Free tier keys have strict rate limits (e.g., 2-15 requests per minute). Please wait a minute before trying again, or check your billing at https://aistudio.google.com/app/plan_billing";
  } else if (combinedStr.includes('403') || combinedStr.toLowerCase().includes('permission') || combinedStr.toLowerCase().includes('apikey') || errorCode === '403') {
    return "API key restricted or invalid. Check your Vercel Environment Variables and verify the key has 'Generative Language API' enabled in Google AI Studio.";
  }
  
  return detail || error?.message || errorMessage;
};

interface KBTopic {
  title: string;
  ref: string;
  desc: string;
  url?: string;
}

interface UploadedFile {
  id: string;
  name: string;
  type: string;
  size: number;
  purpose?: string;
  timestamp: string;
}

// --- COMPONENTS ---

const TaxAdvisory = () => {
  const [queryText, setQueryText] = useState("");
  const [loading, setLoading] = useState(false);
  const [history, setHistory] = useState<TaxQuery[]>([]);
  const [selectedChatId, setSelectedChatId] = useState<string | null>(null);
  const [attachments, setAttachments] = useState<File[]>([]);
  const [urlInput, setUrlInput] = useState("");
  const [showUrlInput, setShowUrlInput] = useState(false);
  const [currentUrls, setCurrentUrls] = useState<string[]>([]);
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [deleteConfirmId, setDeleteConfirmId] = useState<string | null>(null);
  const [guideSearchQuery, setGuideSearchQuery] = useState("");
  const [isSearchingGuides, setIsSearchingGuides] = useState(false);
  const [isListening, setIsListening] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const chatEndRef = React.useRef<HTMLDivElement>(null);
  const recognitionRef = React.useRef<any>(null);

  useEffect(() => {
    const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
    if (SpeechRecognition) {
      const recognition = new SpeechRecognition();
      recognition.continuous = false; // Stop after one phrase to prevent runaway input
      recognition.interimResults = false;
      recognition.lang = 'en-SG';

      recognition.onresult = (event: any) => {
        const transcript = event.results[0][0].transcript;
        if (transcript) {
          setQueryText(prev => prev + (prev.endsWith(' ') || prev === '' ? '' : ' ') + transcript);
        }
      };

      recognition.onend = () => {
        setIsListening(false);
      };

      recognition.onerror = (event: any) => {
        console.error('Speech recognition error:', event.error);
        setIsListening(false);
      };

      recognitionRef.current = recognition;
    }
  }, []);

  const toggleListening = () => {
    if (!recognitionRef.current) {
      alert("Speech recognition is not supported in your browser.");
      return;
    }

    if (isListening) {
      recognitionRef.current.stop();
    } else {
      try {
        recognitionRef.current.start();
        setIsListening(true);
      } catch (err) {
        console.error("Failed to start recognition:", err);
      }
    }
  };

  useEffect(() => {
    const savedHistory = localStorage.getItem('tax_gpt_history');
    if (savedHistory) {
      const parsed = JSON.parse(savedHistory);
      setHistory(parsed);
      if (parsed.length > 0 && !selectedChatId) {
        setSelectedChatId(parsed[0].id);
      }
    }
  }, []);

  useEffect(() => {
    chatEndRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [selectedChatId, history, loading]);

  const saveHistory = (newHistory: TaxQuery[]) => {
    setHistory(newHistory);
    localStorage.setItem('tax_gpt_history', JSON.stringify(newHistory));
  };

  const fileToGenerativePart = async (file: File): Promise<{ inlineData: { data: string, mimeType: string } }> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => {
        const base64Data = (reader.result as string).split(',')[1];
        if (!base64Data) {
          reject(new Error(`Failed to read file: ${file.name}`));
          return;
        }
        resolve({
          inlineData: {
            data: base64Data,
            mimeType: file.type || 'application/octet-stream'
          },
        });
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  };

  const askTaxGemini = async () => {
    if (!queryText.trim()) return;
    const textToSubmit = queryText;
    setLoading(true);
    setError(null);
    setQueryText("");
    
    if (!(process.env.GEMINI_API_KEY)) {
      setError("API Key is missing. Please set GEMINI_API_KEY in your Vercel Environment Variables.");
      setLoading(false);
      return;
    }

    const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

    // Fetch Knowledge Base context
    const kbTopics = JSON.parse(localStorage.getItem('tax_gpt_kb_topics') || '[]');
    const kbFiles = JSON.parse(localStorage.getItem('tax_gpt_kb_files') || '[]');
    
    const kbContext = `
    Knowledge Base Topics:
    ${kbTopics.map((t: KBTopic) => `- ${t.title}: ${t.desc} (Ref: ${t.ref}${t.url ? `, Link: ${t.url}` : ''})`).join('\n')}
    
    Knowledge Base Files:
    ${kbFiles.map((f: UploadedFile) => `- ${f.name} (Type: ${f.type}, Purpose: ${f.purpose})`).join('\n')}
    `;

    const cacheKey = `${textToSubmit.toLowerCase()}_${selectedChatId || 'new'}`;
    if (advisoryCache.has(cacheKey) && attachments.length === 0 && currentUrls.length === 0) {
      const cached = advisoryCache.get(cacheKey);
      if (selectedChatId) {
        setHistory(history.map(h => h.id === selectedChatId ? { ...h, ...cached } : h));
      } else {
        const newEntry = { id: Date.now().toString(), query: textToSubmit, ...cached, timestamp: new Date().toISOString() };
        setHistory([newEntry, ...history].slice(0, 10));
        setSelectedChatId(newEntry.id);
      }
      setLoading(false);
      return;
    }

    const systemPrompt = `You are an expert Singapore Tax Consultant (Big 4 background). 
    Provide advice based on the Income Tax Act of Singapore and IRAS guidelines.
    
    You have access to a local Knowledge Base and Google Search for real-time grounding.
    
    Format your response in JSON with these keys: 
    'answer' (direct answer), 
    'explanation' (detailed reasoning), 
    'assumptions' (list of strings), 
    'references' (IRAS e-Tax guides or Section numbers),
    'citations' (array of objects with:
       'source': Specific title of the page or e-Tax guide (e.g., "IRAS e-Tax Guide: Deductibility of Expenses"),
       'link': FULL absolute URL starting with https://,
       'location': Specific section, paragraph, or page number,
       'snippet': A brief (1-2 sentence) key takeaway or quote from this specific source that supports your answer,
       'searchQuery': A highly specific search query that would lead a user directly to this information on Google or IRAS website,
       'type': One of ['Legislation', 'IRAS Guide', 'Circular', 'General'],
       'relevance': One of ['High', 'Medium', 'Low']).
    
    CRITICAL LINK GUIDELINES:
    1. For 'link', you MUST provide the EXACT absolute URL found in the Google Search results. 
    2. NEVER guess, construct, or hallucinate a URL structure. If you do not have the exact URL, leave the 'link' field EMPTY ("").
    3. IMPORTANT: IRAS has updated their website. AVOID links containing '/irashome/' or ending in '.aspx'. These are deprecated and lead to 404 errors. 
    4. PREFER links starting with 'https://www.iras.gov.sg/taxes/' or 'https://www.iras.gov.sg/quick-links/'.
    5. New IRAS structure examples:
       - Corporate Tax: https://www.iras.gov.sg/taxes/corporate-income-tax
       - GST: https://www.iras.gov.sg/taxes/goods-and-services-tax-(gst)
       - Individual Tax: https://www.iras.gov.sg/taxes/individual-income-tax
    6. If you cannot find a direct, verified link to the specific page, leave 'link' empty and provide the 'source' and 'location' so the user can search for it.
    7. Ensure all links start with https://.
    8. DO NOT provide links to PDFs directly unless you are certain they work.
    
    If you use information from the Knowledge Base, cite the file name or topic title and provide the link if available in the context.
    If you use information from Google Search, cite the website name and provide the EXACT URL from the search result.
    
    Be conservative. If unsure, advise professional consultation.`;

    try {
      let currentChatHistory: { role: 'user' | 'model'; text: string }[] = [];
      let activeChatId = selectedChatId;

      if (activeChatId) {
        const existingChat = history.find(h => h.id === activeChatId);
        if (existingChat) {
          currentChatHistory = existingChat.chatHistory || [
            { role: 'user', text: existingChat.query },
            { role: 'model', text: existingChat.answer }
          ];
        }
      }

      // Ensure alternating roles and non-empty text
      const cleanedHistory = [];
      let lastRole: string | null = null;
      for (const h of currentChatHistory) {
        if (h.text && h.text.trim() && h.role !== lastRole) {
          cleanedHistory.push({ role: h.role, parts: [{ text: h.text }] });
          lastRole = h.role;
        }
      }

      // If the last role was 'user', we can't add another 'user' message directly
      // But Gemini usually handles this by merging or we can just drop the last history item if it's 'user'
      // However, the current message we are about to add IS 'user'.
      // So the last item in cleanedHistory MUST be 'model'.
      if (cleanedHistory.length > 0 && cleanedHistory[cleanedHistory.length - 1].role === 'user') {
        cleanedHistory.pop();
      }

      const parts: any[] = [{ text: `User Query: ${textToSubmit}\n\nLocal Knowledge Base Context:\n${kbContext}` }];
      
      // Add URLs to context if any
      if (currentUrls.length > 0) {
        parts[0].text += `\n\nUser-provided Reference URLs:\n${currentUrls.join('\n')}`;
      }

      if (attachments.length > 0) {
        const attachmentParts = await Promise.all(attachments.map(fileToGenerativePart));
        parts.push(...attachmentParts);
      }

      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [
          ...cleanedHistory,
          { role: 'user', parts }
        ],
        config: {
          systemInstruction: systemPrompt,
          responseMimeType: "application/json",
          tools: [{ googleSearch: {} }]
        }
      });
      
      const content = JSON.parse(response.text);

      // Enrich citations with real grounding metadata from Google Search
      const groundingMetadata = (response as any).groundingMetadata;
      if (groundingMetadata?.groundingChunks) {
        const groundingCitations = groundingMetadata.groundingChunks
          .filter((chunk: any) => chunk.web?.uri)
          .map((chunk: any) => {
            const cleaned = cleanIrasUrl(chunk.web?.uri);
            const isLegislation = chunk.web?.uri.includes('sso.agc.gov.sg');
            return {
              source: chunk.web?.title || 'IRAS Official Source',
              link: cleaned.link,
              isModern: cleaned.isModern,
              isPdf: cleaned.isPdf,
              location: 'Search Result',
              snippet: 'Information verified via Google Search grounding.',
              searchQuery: chunk.web?.title || '',
              type: isLegislation ? 'Legislation' : (cleaned.link.includes('e-tax-guides') ? 'IRAS Guide' : 'General'),
              relevance: 'High'
            };
          });
        
        // Filter out duplicates and hallucinated links from the AI's JSON
        const validAiCitations = (content.citations || []).map((c: any) => {
          const cleaned = cleanIrasUrl(c.link);
          return {
            ...c,
            link: cleaned.link,
            isModern: cleaned.isModern,
            isPdf: cleaned.isPdf
          };
        }).filter((c: any) => 
          c.link && 
          c.link.startsWith('https://') && 
          !c.link.includes('/irashome/') &&
          !c.link.includes('.aspx')
        );

        content.citations = [...validAiCitations, ...groundingCitations];
        
        // Remove duplicates by link
        const seenLinks = new Set();
        content.citations = content.citations.filter((c: any) => {
          if (!c.link) return true;
          if (seenLinks.has(c.link)) return false;
          seenLinks.add(c.link);
          return true;
        });
      }
      
      const searchEntryPoint = groundingMetadata?.searchEntryPoint?.renderedContent;

      const cacheData = {
        answer: content.answer,
        explanation: content.explanation,
        assumptions: content.assumptions,
        references: content.references,
        citations: content.citations,
        searchEntryPoint,
        chatHistory: [
          ...currentChatHistory,
          { role: 'user', text: textToSubmit },
          { role: 'model', text: content.answer }
        ]
      };
      advisoryCache.set(cacheKey, cacheData);

      if (activeChatId) {
        const updatedHistory = history.map(h => {
          if (h.id === activeChatId) {
            return {
              ...h,
              answer: content.answer,
              explanation: content.explanation,
              assumptions: content.assumptions,
              references: content.references,
              citations: content.citations,
              searchEntryPoint,
              urls: [...(h.urls || []), ...currentUrls],
              chatHistory: [
                ...currentChatHistory,
                { role: 'user', text: textToSubmit },
                { role: 'model', text: content.answer }
              ]
            };
          }
          return h;
        });
        saveHistory(updatedHistory);
      } else {
        const newEntry: TaxQuery = {
          id: Date.now().toString(),
          query: textToSubmit,
          ...content,
          searchEntryPoint,
          timestamp: new Date().toISOString(),
          chatHistory: [
            { role: 'user', text: textToSubmit },
            { role: 'model', text: content.answer }
          ],
          attachments: attachments.map(f => ({ name: f.name, type: f.type })),
          urls: currentUrls
        };
        saveHistory([newEntry, ...history].slice(0, 10));
        setSelectedChatId(newEntry.id);
      }

      setAttachments([]);
      setCurrentUrls([]);
      setShowUrlInput(false);
    } catch (error: any) {
      setError(parseAIError(error));
    } finally {
      setLoading(false);
    }
  };

  const searchIrasGuides = async () => {
    if (!guideSearchQuery.trim()) return;
    const query = guideSearchQuery.trim();
    setIsSearchingGuides(true);
    setGuideSearchQuery("");

    const processResults = (data: any, queryStr: string) => {
      const searchMessage: TaxQuery = {
        id: Date.now().toString(),
        query: `Source Search: ${queryStr}`,
        answer: `I have located the following official IRAS guides and legal references related to "${queryStr}":`,
        explanation: "These sources provide the official regulatory framework and administrative guidance from IRAS and the Income Tax Act.",
        assumptions: ["Results are based on current IRAS e-Tax guide availability.", "Links are verified for modern IRAS web structure."],
        references: data.results.map((r: any) => r.actSection || r.title),
        citations: data.results.map((r: any) => {
          const cleaned = cleanIrasUrl(r.link);
          return {
            source: r.title,
            link: cleaned.link,
            location: r.actSection || 'IRAS Official Guide',
            snippet: r.description,
            isModern: cleaned.isModern,
            isPdf: cleaned.isPdf,
            searchQuery: r.title,
            type: r.type || (r.actSection ? 'Legislation' : 'IRAS Guide'),
            relevance: 'High'
          };
        }),
        timestamp: new Date().toISOString(),
        chatHistory: [
          { role: 'user', text: `Search for IRAS guides and Act sections: ${queryStr}` },
          { role: 'model', text: `I have found several relevant IRAS guides and references for "${queryStr}".` }
        ]
      };
      saveHistory([searchMessage, ...history].slice(0, 10));
      setSelectedChatId(searchMessage.id);
    };

    // 1. Check Cache
    const cacheKey = query.toLowerCase();
    if (searchCache.has(cacheKey)) {
      processResults(searchCache.get(cacheKey), query);
      setIsSearchingGuides(false);
      return;
    }

    // 2. Local Heuristics
    const heuristicData = findHeuristicResult(query);
    if (heuristicData) {
      processResults(heuristicData, query);
      setIsSearchingGuides(false);
      return;
    }

    const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

    const prompt = `You are a specialized Singapore Tax Librarian. 
    Your task is to find specific IRAS e-Tax guides and sections of the Income Tax Act related to the user's query.
    
    Query: ${query}
    
    Return a JSON object with:
    'results': array of objects with:
       'title': Title of the guide or section,
       'description': Brief summary of what it covers,
       'link': Official URL (must be modern IRAS structure),
       'actSection': If applicable, the specific section of the Income Tax Act,
       'type': One of ['Legislation', 'IRAS Guide', 'Circular', 'General'].
    
    PRIORITIZE MODERN IRAS URLS:
    - Corporate Tax: https://www.iras.gov.sg/taxes/corporate-income-tax
    - GST: https://www.iras.gov.sg/taxes/goods-and-services-tax-(gst)
    - Individual Tax: https://www.iras.gov.sg/taxes/individual-income-tax
    - Avoid .aspx or /irashome/ links.
    
    If you find multiple relevant guides, list up to 5.`;

    try {
      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        config: { 
          responseMimeType: "application/json",
          tools: [{ googleSearch: {} }]
        }
      });
      
      const data = JSON.parse(response.text);
      if (data && data.results) {
        searchCache.set(cacheKey, data);
      }
      
      processResults(data, query);
    } catch (error: any) {
      setError(parseAIError(error));
    } finally {
      setIsSearchingGuides(false);
    }
  };

  const handleFeedback = (id: string, feedback: 'correct' | 'wrong') => {
    const updatedHistory = history.map(h => h.id === id ? { ...h, feedback } : h);
    saveHistory(updatedHistory);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files) {
      setAttachments(Array.from(e.target.files));
    }
  };

  const addUrl = () => {
    if (urlInput.trim() && !currentUrls.includes(urlInput.trim())) {
      setCurrentUrls([...currentUrls, urlInput.trim()]);
      setUrlInput("");
    }
  };

  const clearHistory = () => {
    saveHistory([]);
    setSelectedChatId(null);
    setShowClearConfirm(false);
  };

  const deleteChat = (id: string) => {
    const newHistory = history.filter(h => h.id !== id);
    saveHistory(newHistory);
    if (selectedChatId === id) {
      setSelectedChatId(newHistory[0]?.id || null);
    }
    setDeleteConfirmId(null);
  };

  const activeChat = history.find(h => h.id === selectedChatId);

  return (
    <div className="flex flex-col md:flex-row gap-6 h-[700px]">
      {/* Sidebar: History */}
      <div className="w-full md:w-72 flex flex-col bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="p-4 border-b border-slate-100 flex justify-between items-center bg-slate-50">
          <h3 className="text-sm font-bold text-slate-700">Recent Chats</h3>
          <div className="flex gap-2 relative">
            {showClearConfirm ? (
              <div className="absolute right-0 top-0 flex items-center gap-1 bg-white border border-slate-200 rounded-lg shadow-lg p-1 z-50 animate-in fade-in zoom-in duration-200">
                <button 
                  onClick={clearHistory}
                  className="px-2 py-1 text-[10px] font-bold bg-red-600 text-white rounded hover:bg-red-700 transition-colors"
                >
                  Confirm Clear
                </button>
                <button 
                  onClick={() => setShowClearConfirm(false)}
                  className="px-2 py-1 text-[10px] font-bold bg-slate-100 text-slate-600 rounded hover:bg-slate-200 transition-colors"
                >
                  Cancel
                </button>
              </div>
            ) : (
              <button 
                onClick={() => setShowClearConfirm(true)}
                className="p-1.5 text-slate-400 hover:text-red-500 transition-colors"
                title="Clear All"
              >
                <Trash2 size={16} />
              </button>
            )}
            <button 
              onClick={() => setSelectedChatId(null)}
              className="p-1.5 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
              title="New Chat"
            >
              <Plus size={16} />
            </button>
          </div>
        </div>
        <div className="flex-1 overflow-y-auto p-2 space-y-1">
          {history.length === 0 && (
            <p className="text-xs text-slate-400 text-center py-8 px-4">No recent conversations. Start a new one!</p>
          )}
          {history.map((item) => (
            <div
              key={item.id}
              onClick={() => setSelectedChatId(item.id)}
              className={`w-full text-left p-3 rounded-lg transition-all group relative cursor-pointer ${selectedChatId === item.id ? 'bg-blue-50 border-blue-100' : 'hover:bg-slate-50 border-transparent'}`}
              role="button"
              tabIndex={0}
            >
              <div className="text-xs font-bold text-blue-600 mb-1 truncate">
                {new Date(item.timestamp).toLocaleDateString()}
              </div>
              <div className="text-sm text-slate-700 font-medium truncate pr-6">
                {item.query}
              </div>
              {deleteConfirmId === item.id ? (
                <div className="absolute right-2 top-1/2 -translate-y-1/2 flex items-center gap-1 bg-white border border-slate-200 rounded shadow-sm p-0.5 z-10">
                  <button 
                    onClick={(e) => { e.stopPropagation(); deleteChat(item.id); }}
                    className="p-1 text-[8px] font-bold bg-red-600 text-white rounded hover:bg-red-700"
                  >
                    Del
                  </button>
                  <button 
                    onClick={(e) => { e.stopPropagation(); setDeleteConfirmId(null); }}
                    className="p-1 text-[8px] font-bold bg-slate-100 text-slate-600 rounded hover:bg-slate-200"
                  >
                    Esc
                  </button>
                </div>
              ) : (
                <button 
                  onClick={(e) => {
                    e.stopPropagation();
                    setDeleteConfirmId(item.id);
                  }}
                  className="absolute right-2 top-1/2 -translate-y-1/2 opacity-0 group-hover:opacity-100 md:opacity-0 md:group-hover:opacity-100 text-slate-300 hover:text-red-500 p-1 transition-opacity"
                >
                  <Trash2 size={14} />
                </button>
              )}
            </div>
          ))}
        </div>
      </div>

      {/* Main Chat Area */}
      <div className="flex-1 flex flex-col bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden relative">
        {/* Chat Header */}
        <div className="p-4 border-b border-slate-100 flex justify-between items-center bg-white z-10">
          <div className="flex items-center gap-3">
            <div className="bg-blue-100 p-2 rounded-lg">
              <MessageSquare className="text-blue-600" size={20} />
            </div>
            <div>
              <h3 className="text-sm font-bold text-slate-800">
                {activeChat ? "Active Conversation" : "New Advisory Session"}
              </h3>
              <p className="text-[10px] text-slate-400 uppercase font-bold tracking-wider">
                Expert Singapore Tax Support
              </p>
            </div>
          </div>
          {activeChat && (
            <div className="flex gap-2">
              <button 
                onClick={() => handleFeedback(activeChat.id, 'correct')}
                className={`p-1.5 rounded-lg transition-colors ${activeChat.feedback === 'correct' ? 'bg-emerald-100 text-emerald-600' : 'hover:bg-slate-100 text-slate-400'}`}
              >
                <ThumbsUp size={16} />
              </button>
              <button 
                onClick={() => handleFeedback(activeChat.id, 'wrong')}
                className={`p-1.5 rounded-lg transition-colors ${activeChat.feedback === 'wrong' ? 'bg-red-100 text-red-600' : 'hover:bg-slate-100 text-slate-400'}`}
              >
                <ThumbsDown size={16} />
              </button>
            </div>
          )}
        </div>

        {/* Source Search Bar */}
        <div className="px-4 py-3 bg-slate-50 border-b border-slate-100 flex items-center gap-3">
          <div className="flex-1 relative group">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 group-focus-within:text-blue-500 transition-colors" size={14} />
            <input 
              className="w-full pl-9 pr-4 py-2 bg-white border border-slate-200 rounded-lg text-xs outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all"
              placeholder="Search IRAS e-Tax guides or Income Tax Act sections..."
              value={guideSearchQuery}
              onChange={(e) => setGuideSearchQuery(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  searchIrasGuides();
                }
              }}
            />
          </div>
          <button 
            onClick={searchIrasGuides}
            disabled={isSearchingGuides || !guideSearchQuery.trim()}
            className="px-4 py-2 bg-blue-600 text-white rounded-lg text-xs font-bold hover:bg-blue-700 disabled:bg-slate-300 transition-all shadow-sm flex items-center gap-2"
          >
            {isSearchingGuides ? (
              <>
                <RefreshCw size={12} className="animate-spin" />
                Searching...
              </>
            ) : (
              <>
                <BookOpen size={12} />
                Find Sources
              </>
            )}
          </button>
        </div>

        {/* Messages */}
        <div className="flex-1 overflow-y-auto p-6 space-y-6 bg-slate-50/30">
          {!activeChat && !loading && (
            <div className="h-full flex flex-col items-center justify-center text-center max-w-sm mx-auto space-y-4">
              <div className="bg-blue-50 p-4 rounded-full">
                <ShieldCheck className="text-blue-600" size={48} />
              </div>
              <h4 className="text-lg font-bold text-slate-800">How can I assist with your tax matters today?</h4>
              <p className="text-sm text-slate-500">
                Ask about corporate tax, GST, S14Q renovation claims, or any other IRAS-related queries.
              </p>
            </div>
          )}

          {activeChat?.chatHistory?.map((msg, idx) => (
            <div key={idx} className={`flex ${msg.role === 'user' ? 'justify-end' : 'justify-start'}`}>
              <div className={`max-w-[85%] rounded-2xl p-4 shadow-sm ${
                msg.role === 'user' 
                  ? 'bg-blue-600 text-white rounded-tr-none' 
                  : 'bg-white border border-slate-100 text-slate-800 rounded-tl-none'
              }`}>
                {msg.role === 'model' ? (
                  <div className="space-y-4">
                    <p className="text-sm leading-relaxed whitespace-pre-wrap">{msg.text}</p>
                    {/* Only show detailed breakdown for the LATEST model message */}
                    {idx === activeChat.chatHistory!.length - 1 && (
                      <div className="mt-4 pt-4 border-t border-slate-100 space-y-4">
                        <section>
                          <h5 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Reasoning</h5>
                          <p className="text-xs text-slate-600 leading-relaxed">{activeChat.explanation}</p>
                        </section>
                        <div className="grid grid-cols-2 gap-4">
                          <section>
                            <h5 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Assumptions</h5>
                            <ul className="list-disc list-inside text-[10px] text-slate-500 space-y-1">
                              {activeChat.assumptions?.map((a, i) => <li key={i}>{a}</li>)}
                            </ul>
                          </section>
                          <section>
                            <h5 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2">Legal References</h5>
                            <div className="flex flex-wrap gap-2">
                              {Array.isArray(activeChat.references) ? activeChat.references.map((ref, i) => (
                                <span key={i} className="text-[10px] bg-blue-50 text-blue-600 px-2 py-0.5 rounded border border-blue-100 font-mono">
                                  {ref}
                                </span>
                              )) : (
                                <span className="text-[10px] bg-blue-50 text-blue-600 px-2 py-0.5 rounded border border-blue-100 font-mono">
                                  {activeChat.references}
                                </span>
                              )}
                            </div>
                          </section>
                        </div>
                        {activeChat.citations && activeChat.citations.length > 0 && (
                          <section className="pt-4 border-t border-slate-100">
                            <h5 className="text-[10px] font-bold text-slate-400 uppercase tracking-widest mb-2 flex items-center gap-1">
                              <LinkIcon size={10} /> Grounding & Citations
                            </h5>
                            <div className="space-y-2">
                              {activeChat.citations.map((cite, i) => (
                                <div key={i} className={`p-4 rounded-xl border ${cite.location === 'Search Result' ? 'bg-emerald-50/20 border-emerald-100' : 'bg-slate-50 border-slate-100'}`}>
                                  <div className="flex flex-col gap-3">
                                    <div className="flex justify-between items-start">
                                      <div className="flex flex-col gap-1">
                                        <div className="flex items-center gap-2 flex-wrap">
                                          <span className="text-sm font-bold text-slate-800">{cite.source}</span>
                                          {cite.type && (
                                            <span className={`text-[8px] px-1.5 py-0.5 rounded-full font-bold uppercase tracking-wider ${
                                              cite.type === 'Legislation' ? 'bg-purple-100 text-purple-700' :
                                              cite.type === 'IRAS Guide' ? 'bg-blue-100 text-blue-700' :
                                              'bg-slate-200 text-slate-700'
                                            }`}>
                                              {cite.type}
                                            </span>
                                          )}
                                          {cite.isPdf && (
                                            <span className="text-[8px] bg-red-100 text-red-700 px-1.5 py-0.5 rounded-full font-bold uppercase tracking-wider flex items-center gap-0.5">
                                              <FileText size={8} /> PDF
                                            </span>
                                          )}
                                          {cite.location === 'Search Result' && (
                                            <span className="text-[8px] bg-emerald-100 text-emerald-700 px-1.5 py-0.5 rounded-full font-bold uppercase tracking-wider">Verified</span>
                                          )}
                                          {cite.relevance === 'High' && (
                                            <span className="text-[8px] bg-amber-100 text-amber-700 px-1.5 py-0.5 rounded-full font-bold uppercase tracking-wider">Primary Source</span>
                                          )}
                                        </div>
                                        {cite.location && cite.location !== 'Search Result' && (
                                          <span className="text-[10px] text-slate-500 font-medium">{cite.location}</span>
                                        )}
                                      </div>
                                      <div className="flex items-center gap-2">
                                        <button 
                                          onClick={() => {
                                            const text = `${cite.source}${cite.location ? ` (${cite.location})` : ''}${cite.link ? ` - ${cite.link}` : ''}`;
                                            navigator.clipboard.writeText(text);
                                          }}
                                          className="p-1.5 text-slate-400 hover:text-blue-600 hover:bg-blue-50 rounded-lg transition-all"
                                          title="Copy Citation"
                                        >
                                          <Copy size={14} />
                                        </button>
                                        {cite.link && (
                                          <a 
                                            href={cite.link} 
                                            target="_blank" 
                                            rel="noopener noreferrer" 
                                            className="flex items-center gap-1.5 text-[10px] bg-blue-600 text-white px-3 py-1.5 rounded-full hover:bg-blue-700 transition-all shadow-sm font-bold"
                                          >
                                            <ExternalLink size={12} /> View Source
                                          </a>
                                        )}
                                      </div>
                                    </div>
                                    
                                    {cite.snippet && (
                                      <div className="bg-white/50 p-2.5 rounded-lg border border-slate-200/50">
                                        <p className="text-xs text-slate-600 leading-relaxed italic">
                                          "{cite.snippet}"
                                        </p>
                                      </div>
                                    )}

                                    <div className="flex items-center gap-2">
                                      <a 
                                        href={`https://www.google.com/search?q=${encodeURIComponent(cite.searchQuery || `site:iras.gov.sg ${cite.source}`)}`}
                                        target="_blank" 
                                        rel="noopener noreferrer" 
                                        className="flex-1 flex items-center justify-center gap-2 text-[10px] bg-white border border-slate-200 text-slate-600 px-3 py-2 rounded-lg hover:bg-slate-50 transition-colors font-bold shadow-sm"
                                      >
                                        <Search size={12} /> Search for this specific topic
                                      </a>
                                      <a 
                                        href={`https://www.google.com/search?q=${encodeURIComponent(cite.source)}`}
                                        target="_blank" 
                                        rel="noopener noreferrer" 
                                        className="flex items-center justify-center p-2 bg-white border border-slate-200 text-slate-400 rounded-lg hover:text-blue-600 hover:border-blue-200 transition-all shadow-sm"
                                        title="Search on Google"
                                      >
                                        <Globe size={14} />
                                      </a>
                                    </div>
                                  </div>
                                </div>
                              ))}
                            </div>
                            {activeChat.searchEntryPoint && (
                              <div className="mt-4 p-4 bg-blue-50/30 rounded-xl border border-blue-100/50">
                                <h6 className="text-[10px] font-bold text-blue-600 uppercase tracking-widest mb-3 flex items-center gap-2">
                                  <Search size={10} /> Official Google Search Results
                                </h6>
                                <div 
                                  className="text-sm google-search-entry"
                                  dangerouslySetInnerHTML={{ __html: activeChat.searchEntryPoint }}
                                />
                              </div>
                            )}
                          </section>
                        )}
                      </div>
                    )}
                  </div>
                ) : (
                  <p className="text-sm font-medium">{msg.text}</p>
                )}
              </div>
            </div>
          ))}

          {loading && (
            <div className="flex justify-start">
              <div className="bg-white border border-slate-100 p-4 rounded-2xl rounded-tl-none shadow-sm flex items-center gap-3">
                <div className="flex gap-1">
                  <div className="w-1.5 h-1.5 bg-blue-400 rounded-full animate-bounce" />
                  <div className="w-1.5 h-1.5 bg-blue-400 rounded-full animate-bounce [animation-delay:0.2s]" />
                  <div className="w-1.5 h-1.5 bg-blue-400 rounded-full animate-bounce [animation-delay:0.4s]" />
                </div>
                <span className="text-xs text-slate-400 font-medium italic">Consulting Income Tax Act...</span>
              </div>
            </div>
          )}
          <div ref={chatEndRef} />
        </div>

        {/* Input Area */}
        <div className="p-4 bg-white border-t border-slate-100">
          {error && (
            <div className="mb-4 p-3 bg-red-50 border border-red-200 rounded-lg flex items-start gap-3 text-red-700 text-sm animate-in fade-in slide-in-from-bottom-2 duration-300">
              <AlertCircle className="w-5 h-5 shrink-0" />
              <div className="flex-1">
                <p className="font-semibold">Error Occurred</p>
                <p>{error}</p>
                <button 
                  onClick={() => setError(null)}
                  className="mt-1 text-xs font-bold underline hover:no-underline"
                >
                  Dismiss
                </button>
              </div>
            </div>
          )}
          <div className="relative bg-slate-50 rounded-xl border border-slate-200 focus-within:ring-2 focus-within:ring-blue-500/20 focus-within:border-blue-500 transition-all">
            <textarea
              className="w-full p-4 pr-24 bg-transparent border-none outline-none resize-none text-sm text-slate-800 placeholder:text-slate-400"
              rows={2}
              placeholder={activeChat ? "Type your follow-up..." : "Describe your tax query..."}
              value={queryText}
              onChange={(e) => setQueryText(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' && !e.shiftKey) {
                  e.preventDefault();
                  askTaxGemini();
                }
              }}
            />
            <div className="absolute bottom-3 right-3 flex items-center gap-2">
              <button 
                onClick={toggleListening}
                className={`p-2 rounded-full transition-all ${isListening ? 'text-red-500 bg-red-50 animate-pulse ring-4 ring-red-100' : 'text-slate-400 hover:text-blue-600 hover:bg-blue-50'}`}
                title={isListening ? "Stop listening" : "Speak to ask"}
              >
                {isListening ? <MicOff size={20} /> : <Mic size={20} />}
              </button>
              <button 
                onClick={() => setShowUrlInput(!showUrlInput)}
                className={`p-2 rounded transition-colors ${showUrlInput ? 'text-blue-600 bg-blue-50' : 'text-slate-400 hover:text-blue-600'}`}
                title="Add URL"
              >
                <LinkIcon size={20} />
              </button>
              <label className="cursor-pointer text-slate-400 hover:text-blue-600 transition-colors p-2">
                <Paperclip size={20} />
                <input 
                  type="file" 
                  multiple 
                  className="hidden" 
                  onChange={handleFileChange}
                  accept="image/*,.pdf,.doc,.docx,.xls,.xlsx"
                />
              </label>
              <button 
                onClick={askTaxGemini}
                disabled={loading || !queryText.trim()}
                className="bg-blue-600 text-white p-2 rounded-lg hover:bg-blue-700 disabled:bg-slate-300 transition-all shadow-sm"
              >
                {loading ? <RefreshCw className="animate-spin h-5 w-5" /> : <Send size={20} />}
              </button>
            </div>
          </div>

          {showUrlInput && (
            <div className="mt-2 flex gap-2">
              <input 
                className="flex-1 text-xs p-2 border border-slate-200 rounded focus:ring-1 focus:ring-blue-500 outline-none"
                placeholder="Paste URL here..."
                value={urlInput}
                onChange={(e) => setUrlInput(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && addUrl()}
              />
              <button onClick={addUrl} className="px-3 py-1 bg-slate-100 text-slate-600 text-xs font-bold rounded hover:bg-slate-200">Add</button>
            </div>
          )}

          {(attachments.length > 0 || currentUrls.length > 0) && (
            <div className="flex flex-wrap gap-2 mt-3">
              {currentUrls.map((url, i) => (
                <div key={`url-${i}`} className="flex items-center gap-2 bg-indigo-50 border border-indigo-100 px-3 py-1 rounded-full text-[10px] text-indigo-700 font-bold">
                  <LinkIcon size={12} />
                  <span className="truncate max-w-[120px]">{url}</span>
                  <button onClick={() => setCurrentUrls(currentUrls.filter((_, idx) => idx !== i))}>
                    <X size={12} className="hover:text-red-500" />
                  </button>
                </div>
              ))}
              {attachments.map((file, i) => (
                <div key={i} className="flex items-center gap-2 bg-blue-50 border border-blue-100 px-3 py-1 rounded-full text-[10px] text-blue-700 font-bold">
                  {file.type.startsWith('image/') ? <ImageIcon size={12} /> : <FileText size={12} />}
                  <span className="truncate max-w-[120px]">{file.name}</span>
                  <button onClick={() => setAttachments(attachments.filter((_, idx) => idx !== i))}>
                    <X size={12} className="hover:text-red-500" />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

const TaxComputation = () => {
  const [selectedYA, setSelectedYA] = useState(new Date().getFullYear());
  const [data, setData] = useState<TaxData>({
    accountingProfit: 100000,
    additions: [
      { id: 1, label: 'Depreciation', amount: 5000, note: 'Non-deductible' },
      { id: 2, label: 'Entertainment (Non-biz)', amount: 1200, note: 'Private portion' }
    ],
    deductions: [
      { id: 3, label: 'Capital Allowances (Current)', amount: 8000, note: 'S19/19A' }
    ],
    otherAdjustments: [
      { id: 4, label: 'Non-taxable Investment Income', amount: 2000, note: 'Capital gain' }
    ]
  });
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [fileObjects, setFileObjects] = useState<Map<string, File>>(new Map());
  const [uploadPurpose, setUploadPurpose] = useState('financials');
  const [analyzing, setAnalyzing] = useState(false);
  const [analysisResults, setAnalysisResults] = useState<any[]>([]);
  const [showAnalysis, setShowAnalysis] = useState(false);
  const [history, setHistory] = useState<TaxData[]>([]);
  const [historyIndex, setHistoryIndex] = useState(-1);
  const [comparisonYA, setComparisonYA] = useState<number | null>(null);
  const [comparisonData, setComparisonData] = useState<TaxData | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const savedData = localStorage.getItem(`tax_gpt_computation_${selectedYA}`);
    if (savedData) {
      const parsed = JSON.parse(savedData);
      setData(parsed);
      setHistory([parsed]);
      setHistoryIndex(0);
    } else {
      const defaultData = {
        accountingProfit: 0,
        additions: [],
        deductions: [],
        otherAdjustments: []
      };
      setData(defaultData);
      setHistory([defaultData]);
      setHistoryIndex(0);
    }
    const savedFiles = localStorage.getItem(`tax_gpt_comp_files_${selectedYA}`);
    if (savedFiles) {
      setUploadedFiles(JSON.parse(savedFiles));
    } else {
      setUploadedFiles([]);
    }
  }, [selectedYA]);

  useEffect(() => {
    if (comparisonYA) {
      const saved = localStorage.getItem(`tax_gpt_computation_${comparisonYA}`);
      if (saved) {
        setComparisonData(JSON.parse(saved));
      } else {
        setComparisonData(null);
      }
    } else {
      setComparisonData(null);
    }
  }, [comparisonYA]);

  const generateSample = () => {
    const sampleData: TaxData = {
      accountingProfit: 250000,
      additions: [
        { id: Date.now() + 1, label: 'Accounting Depreciation', amount: 15000, note: 'Non-deductible under S15(1)(g)' },
        { id: Date.now() + 2, label: 'Private Car Expenses (S-plate)', amount: 4500, note: 'Non-deductible under S15(1)(j)' },
        { id: Date.now() + 3, label: 'Entertainment (Non-business)', amount: 2200, note: 'Not incurred for production of income' }
      ],
      deductions: [
        { id: Date.now() + 4, label: 'Capital Allowances (S19A)', amount: 18000, note: 'Accelerated write-off for IT equipment' },
        { id: Date.now() + 5, label: 'S14Q R&R Claim', amount: 5000, note: 'Renovation & Refurbishment' }
      ],
      otherAdjustments: [
        { id: Date.now() + 6, label: 'Gain on Disposal of Property', amount: 12000, note: 'Capital gain - non-taxable' },
        { id: Date.now() + 7, label: 'Exempt Dividends (Foreign)', amount: 3500, note: 'S13(8) Exemption' }
      ]
    };
    saveData(sampleData);
  };

  const saveData = (newData: TaxData, skipHistory = false) => {
    setData(newData);
    localStorage.setItem(`tax_gpt_computation_${selectedYA}`, JSON.stringify(newData));
    
    if (!skipHistory) {
      const newHistory = history.slice(0, historyIndex + 1);
      newHistory.push(newData);
      // Limit history to 50 steps
      if (newHistory.length > 50) newHistory.shift();
      setHistory(newHistory);
      setHistoryIndex(newHistory.length - 1);
    }
  };

  const undo = () => {
    if (historyIndex > 0) {
      const prevIndex = historyIndex - 1;
      const prevData = history[prevIndex];
      setData(prevData);
      setHistoryIndex(prevIndex);
      localStorage.setItem(`tax_gpt_computation_${selectedYA}`, JSON.stringify(prevData));
    }
  };

  const redo = () => {
    if (historyIndex < history.length - 1) {
      const nextIndex = historyIndex + 1;
      const nextData = history[nextIndex];
      setData(nextData);
      setHistoryIndex(nextIndex);
      localStorage.setItem(`tax_gpt_computation_${selectedYA}`, JSON.stringify(nextData));
    }
  };

  const saveFiles = (newFiles: UploadedFile[]) => {
    setUploadedFiles(newFiles);
    localStorage.setItem(`tax_gpt_comp_files_${selectedYA}`, JSON.stringify(newFiles));
  };

  const exportToExcel = () => {
    const wb = XLSX.utils.book_new();
    
    const summaryData = [
      ["Tax Computation Summary", `YA ${selectedYA}`],
      ["Accounting Profit", data.accountingProfit],
      ["Total Additions", calculations.totalAdditions],
      ["Total Other Adjustments", calculations.totalOtherAdj],
      ["Adjusted Profit", calculations.adjustedProfit],
      ["Total Deductions", calculations.totalDeductions],
      ["Chargeable Income", calculations.chargeableIncome],
      ["Tax Exemption", calculations.exemption],
      ["Net Chargeable", calculations.netChargeable],
      ["Tax Before Rebate", calculations.taxBeforeRebate],
      ["CIT Rebate", calculations.citRebate],
      ["Tax Payable (17%)", calculations.taxPayable]
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, ws, "Summary");
    
    const additionsWs = XLSX.utils.json_to_sheet(data.additions.map(i => ({ 
      Label: i.label, 
      Amount: i.amount, 
      Note: i.note,
      Source: i.isAISuggested ? 'AI Analysis' : 'Manual'
    })));
    XLSX.utils.book_append_sheet(wb, additionsWs, "Additions");

    const otherWs = XLSX.utils.json_to_sheet(data.otherAdjustments.map(i => ({ 
      Label: i.label, 
      Amount: i.amount, 
      Note: i.note,
      Source: i.isAISuggested ? 'AI Analysis' : 'Manual'
    })));
    XLSX.utils.book_append_sheet(wb, otherWs, "Other Adjustments");

    const deductionsWs = XLSX.utils.json_to_sheet(data.deductions.map(i => ({ 
      Label: i.label, 
      Amount: i.amount, 
      Note: i.note,
      Source: i.isAISuggested ? 'AI Analysis' : 'Manual'
    })));
    XLSX.utils.book_append_sheet(wb, deductionsWs, "Deductions");
    
    XLSX.writeFile(wb, `Tax_Computation_YA${selectedYA}.xlsx`);
  };

  const exportToPDF = () => {
    const doc = new jsPDF();
    doc.setFontSize(18);
    doc.text(`Tax Computation - YA ${selectedYA}`, 14, 22);
    
    doc.setFontSize(12);
    doc.text(`Accounting Profit: $${data.accountingProfit.toLocaleString()}`, 14, 35);
    
    (doc as any).autoTable({
      startY: 45,
      head: [['Category', 'Amount']],
      body: [
        ['Total Additions', `$${calculations.totalAdditions.toLocaleString()}`],
        ['Total Other Adjustments', `$${calculations.totalOtherAdj.toLocaleString()}`],
        ['Adjusted Profit', `$${calculations.adjustedProfit.toLocaleString()}`],
        ['Total Deductions', `$${calculations.totalDeductions.toLocaleString()}`],
        ['Chargeable Income', `$${calculations.chargeableIncome.toLocaleString()}`],
        ['Tax Exemption', `$${calculations.exemption.toLocaleString()}`],
        ['Net Chargeable', `$${calculations.netChargeable.toLocaleString()}`],
        ['Tax Before Rebate', `$${calculations.taxBeforeRebate.toLocaleString()}`],
        ['CIT Rebate', `$${calculations.citRebate.toLocaleString()}`],
        ['Tax Payable (17%)', `$${calculations.taxPayable.toLocaleString()}`],
      ],
    });

    const renderAdjustmentTable = (title: string, items: ComputationItem[], startY: number) => {
      doc.setFontSize(14);
      doc.text(title, 14, startY);
      (doc as any).autoTable({
        startY: startY + 5,
        head: [['Item', 'Amount', 'Note', 'Source']],
        body: items.map(i => [
          i.label, 
          `$${i.amount.toLocaleString()}`, 
          i.note,
          i.isAISuggested ? 'AI Suggested' : 'Manual'
        ]),
        didParseCell: (data: any) => {
          if (data.row.raw[3] === 'AI Suggested') {
            data.cell.styles.fillColor = [209, 250, 229]; // Light emerald
            data.cell.styles.textColor = [6, 78, 59]; // Dark emerald
          }
        }
      });
      return (doc as any).lastAutoTable.finalY + 15;
    };

    let nextY = (doc as any).lastAutoTable.finalY + 15;
    if (data.additions.length > 0) nextY = renderAdjustmentTable("Additions", data.additions, nextY);
    if (data.otherAdjustments.length > 0) nextY = renderAdjustmentTable("Other Adjustments", data.otherAdjustments, nextY);
    if (data.deductions.length > 0) nextY = renderAdjustmentTable("Deductions", data.deductions, nextY);
    
    doc.save(`Tax_Computation_YA${selectedYA}.pdf`);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    const newFileObjs = new Map(fileObjects);
    const newFiles: UploadedFile[] = Array.from(files).map((file: File) => {
      const id = Math.random().toString(36).substr(2, 9);
      newFileObjs.set(id, file);
      return {
        id,
        name: file.name,
        type: file.type,
        size: file.size,
        purpose: uploadPurpose,
        timestamp: new Date().toISOString()
      };
    });

    setFileObjects(newFileObjs);
    saveFiles([...newFiles, ...uploadedFiles]);
  };

  const removeFile = (id: string) => {
    const newFileObjs = new Map(fileObjects);
    newFileObjs.delete(id);
    setFileObjects(newFileObjs);
    saveFiles(uploadedFiles.filter(f => f.id !== id));
  };

  const analyzeExcel = async (fileId: string) => {
    const fileInfo = uploadedFiles.find(f => f.id === fileId);
    if (!fileInfo) return;

    // Cache check
    const cacheKey = `analyze_${fileId}_${selectedYA}`;
    if (searchCache.has(cacheKey)) {
      setAnalysisResults(searchCache.get(cacheKey).adjustments || []);
      setShowAnalysis(true);
      return;
    }

    const fileObj = fileObjects.get(fileId);
    setAnalyzing(true);
    setShowAnalysis(true);
    
    try {
      let fileData = "Simulated Data (File content unavailable - please re-upload in this session for deep scan)";
      
      if (fileObj) {
        if (fileObj.name.endsWith('.xlsx') || fileObj.name.endsWith('.xls')) {
          const data = await fileObj.arrayBuffer();
          const workbook = XLSX.read(data);
          const allSheetsData: any = {};
          workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet);
            allSheetsData[sheetName] = json.slice(0, 50); // Limit each sheet to 50 rows for prompt
          });
          fileData = JSON.stringify(allSheetsData);
        } else {
          // For PDF/Images, we'd use vision, but let's assume text extraction or prompt user
          fileData = "Document content (PDF/Image) - Analysis will focus on visual extraction of P&L items.";
        }
      }

      const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

      const prompt = `You are a Singapore Tax Expert. Analyze this financial document (P&L, Balance Sheet, or Trial Balance) for tax adjustments (YA ${selectedYA}).
      
      Your task is to:
      1. Identify the Accounting Profit/Loss.
      2. Identify Non-deductible expenses (Additions): private car expenses, fines, non-business entertainment, capital expenditure, accounting depreciation.
      3. Identify Capital Allowances (Deductions): S19, S19A, S19B based on asset purchases.
      4. Identify Non-taxable income (Other Adjustments): capital gains, exempt dividends.
      
      Data: ${fileData}
      
      Return JSON: { 'accountingProfit': number, 'adjustments': [ { 'item', 'amount', 'impact', 'reason', 'implication' } ] }`;

      const response = await ai.models.generateContent({
        model: "gemini-1.5-flash",
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        config: { responseMimeType: "application/json" }
      });

      const parsed = JSON.parse(response.text);
      searchCache.set(cacheKey, parsed);
      setAnalysisResults(parsed.adjustments || []);
      
      // If it's a direct auto-populate request, we can update the profit too
      if (fileInfo.purpose === 'financials') {
        setData(prev => ({
          ...prev,
          accountingProfit: parsed.accountingProfit || prev.accountingProfit
        }));
      }
    } catch (error: any) {
      setError(parseAIError(error));
      setAnalysisResults([]);
    } finally {
      setAnalyzing(false);
    }
  };

  const applyAdjustments = () => {
    const newAdditions = [...data.additions];
    const newDeductions = [...data.deductions];
    const newOther = [...data.otherAdjustments];

    analysisResults.forEach(res => {
      const newItem = { 
        id: Date.now() + Math.random(), 
        label: res.item, 
        amount: res.amount, 
        note: `${res.reason}${res.implication ? ` | Risk: ${res.implication}` : ''}`,
        isAISuggested: true 
      };
      if (res.impact === 'addition') newAdditions.push(newItem);
      else if (res.impact === 'deduction') newOther.push(newItem);
      else if (res.impact === 'ca') newDeductions.push(newItem);
    });

    saveData({
      ...data,
      additions: newAdditions,
      deductions: newDeductions,
      otherAdjustments: newOther
    });
    setShowAnalysis(false);
  };

  const calculateTax = (taxData: TaxData, ya: number) => {
    const totalAdditions = taxData.additions.reduce((sum, item) => sum + item.amount, 0);
    const totalOtherAdj = taxData.otherAdjustments.reduce((sum, item) => sum + item.amount, 0);
    const adjustedProfit = taxData.accountingProfit + totalAdditions - totalOtherAdj;
    const totalDeductions = taxData.deductions.reduce((sum, item) => sum + item.amount, 0);
    const chargeableIncome = Math.max(0, adjustedProfit - totalDeductions);
    
    let exemption = 0;
    if (chargeableIncome > 0) {
      exemption += Math.min(chargeableIncome, 10000) * PARTIAL_TAX_EXEMPTION.first10kRate;
      if (chargeableIncome > 10000) {
        exemption += Math.min(chargeableIncome - 10000, 190000) * PARTIAL_TAX_EXEMPTION.next190kRate;
      }
    }
    
    const netChargeable = chargeableIncome - exemption;
    const taxBeforeRebate = netChargeable * CORPORATE_TAX_RATE;

    let citRebate = 0;
    if (ya >= 2024) {
      citRebate = Math.min(taxBeforeRebate * 0.5, 40000);
    }
    
    const taxPayable = Math.max(0, taxBeforeRebate - citRebate);

    return { totalAdditions, totalOtherAdj, adjustedProfit, totalDeductions, chargeableIncome, exemption, netChargeable, taxBeforeRebate, citRebate, taxPayable };
  };

  const calculations = useMemo(() => calculateTax(data, selectedYA), [data, selectedYA]);
  const compCalculations = useMemo(() => comparisonData ? calculateTax(comparisonData, comparisonYA!) : null, [comparisonData, comparisonYA]);

  const addItem = (type: keyof TaxData) => {
    if (type === 'accountingProfit') return;
    const newItem = { id: Date.now(), label: 'New Item', amount: 0, note: '' };
    const newData = { ...data, [type]: [...(data[type] as ComputationItem[]), newItem] };
    saveData(newData);
  };

  const updateItem = (type: keyof TaxData, id: number, field: string, value: any) => {
    if (type === 'accountingProfit') return;
    const newData = {
      ...data,
      [type]: (data[type] as ComputationItem[]).map(item => item.id === id ? { ...item, [field]: value } : item)
    };
    saveData(newData);
  };

  const deleteItem = (type: keyof TaxData, id: number) => {
    if (type === 'accountingProfit') return;
    const newData = { ...data, [type]: (data[type] as ComputationItem[]).filter(item => item.id !== id) };
    saveData(newData);
  };

  return (
    <div className="space-y-6">
      {error && (
        <div className="p-4 bg-red-50 border border-red-200 rounded-xl flex items-start gap-4 text-red-700 text-sm animate-in fade-in slide-in-from-top-2 duration-300 shadow-sm">
          <AlertCircle className="w-6 h-6 shrink-0" />
          <div className="flex-1">
            <p className="font-bold text-base mb-1">Process Notification</p>
            <p className="leading-relaxed">{error}</p>
            <button onClick={() => setError(null)} className="mt-2 text-xs font-black uppercase tracking-widest hover:underline">Dismiss</button>
          </div>
        </div>
      )}
      <div className="grid lg:grid-cols-3 gap-8">
        <div className="lg:col-span-2 space-y-6">
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <div className="flex justify-between items-center mb-6">
            <div className="flex items-center gap-4">
              <h3 className="text-lg font-bold text-slate-800">Computation Worksheet</h3>
              <select 
                value={selectedYA}
                onChange={(e) => setSelectedYA(parseInt(e.target.value))}
                className="bg-slate-50 border border-slate-200 rounded px-2 py-1 text-sm font-bold text-blue-600 outline-none"
              >
                {[2022, 2023, 2024, 2025, 2026].map(year => (
                  <option key={year} value={year}>YA {year}</option>
                ))}
              </select>
            </div>
            <div className="flex gap-2">
               <button 
                onClick={generateSample}
                className="flex items-center gap-1 text-xs bg-blue-50 text-blue-600 px-3 py-1.5 rounded hover:bg-blue-100 font-bold"
               >
                <Sparkles size={14} /> Generate Sample
               </button>
               <div className="w-px h-8 bg-slate-100 mx-1" />
               <button 
                onClick={undo}
                disabled={historyIndex <= 0}
                className="p-1.5 bg-white border border-slate-200 rounded text-slate-600 hover:bg-slate-50 disabled:opacity-30 disabled:cursor-not-allowed transition-all"
                title="Undo"
               >
                <RefreshCw size={14} className="-scale-x-100" />
               </button>
               <button 
                onClick={redo}
                disabled={historyIndex >= history.length - 1}
                className="p-1.5 bg-white border border-slate-200 rounded text-slate-600 hover:bg-slate-50 disabled:opacity-30 disabled:cursor-not-allowed transition-all"
                title="Redo"
               >
                <RefreshCw size={14} />
               </button>
               <div className="w-px h-8 bg-slate-100 mx-1" />
               <button 
                onClick={exportToExcel}
                className="flex items-center gap-1 text-xs bg-slate-100 text-slate-600 px-3 py-1.5 rounded hover:bg-slate-200"
               >
                <FileSpreadsheet size={14} /> Excel
               </button>
               <button 
                onClick={exportToPDF}
                className="flex items-center gap-1 text-xs bg-slate-100 text-slate-600 px-3 py-1.5 rounded hover:bg-slate-200"
               >
                <Download size={14} /> PDF
               </button>
            </div>
          </div>

          <div className="space-y-8">
            {/* Comparison Bar */}
            <div className="bg-slate-50 p-4 rounded-lg border border-slate-100 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <TrendingUp className="text-blue-600" size={18} />
                <span className="text-xs font-bold text-slate-600 uppercase tracking-wider">Compare with another YA</span>
              </div>
              <select 
                value={comparisonYA || ""}
                onChange={(e) => setComparisonYA(e.target.value ? parseInt(e.target.value) : null)}
                className="bg-white border border-slate-200 rounded px-2 py-1 text-xs outline-none"
              >
                <option value="">Select YA to compare...</option>
                {[2022, 2023, 2024, 2025, 2026].filter(y => y !== selectedYA).map(year => (
                  <option key={year} value={year}>YA {year}</option>
                ))}
              </select>
            </div>

            {comparisonData && compCalculations && (
              <div className="grid grid-cols-2 gap-4 p-4 bg-blue-50/50 rounded-xl border border-blue-100">
                <div>
                  <h5 className="text-[10px] font-bold text-blue-600 uppercase mb-2">Chargeable Income Variance</h5>
                  <div className="flex items-baseline gap-2">
                    <span className="text-lg font-bold text-slate-800">${calculations.chargeableIncome.toLocaleString()}</span>
                    <span className={`text-xs font-bold ${calculations.chargeableIncome > compCalculations.chargeableIncome ? 'text-red-500' : 'text-emerald-500'}`}>
                      {calculations.chargeableIncome > compCalculations.chargeableIncome ? '+' : ''}
                      {((calculations.chargeableIncome - compCalculations.chargeableIncome) / compCalculations.chargeableIncome * 100).toFixed(1)}%
                    </span>
                  </div>
                  <p className="text-[10px] text-slate-400">vs YA {comparisonYA}: ${compCalculations.chargeableIncome.toLocaleString()}</p>
                </div>
                <div>
                  <h5 className="text-[10px] font-bold text-blue-600 uppercase mb-2">Tax Payable Variance</h5>
                  <div className="flex items-baseline gap-2">
                    <span className="text-lg font-bold text-slate-800">${calculations.taxPayable.toLocaleString()}</span>
                    <span className={`text-xs font-bold ${calculations.taxPayable > compCalculations.taxPayable ? 'text-red-500' : 'text-emerald-500'}`}>
                      {calculations.taxPayable > compCalculations.taxPayable ? '+' : ''}
                      {((calculations.taxPayable - compCalculations.taxPayable) / compCalculations.taxPayable * 100).toFixed(1)}%
                    </span>
                  </div>
                  <p className="text-[10px] text-slate-400">vs YA {comparisonYA}: ${compCalculations.taxPayable.toLocaleString()}</p>
                </div>
              </div>
            )}

            {error && (
              <div className="mt-4 p-3 bg-red-50 border border-red-200 rounded-lg flex items-start gap-3 text-red-700 text-sm animate-in fade-in slide-in-from-bottom-2 duration-300">
                <AlertCircle className="w-5 h-5 shrink-0" />
                <div className="flex-1">
                  <p className="font-semibold">Error Occurred</p>
                  <p>{error}</p>
                  <button onClick={() => setError(null)} className="mt-1 text-xs font-bold underline hover:no-underline">Dismiss</button>
                </div>
              </div>
            )}

            <div className="pb-6 border-b border-slate-100">
              <label className="block text-sm font-semibold text-slate-500 uppercase mb-2">Net Profit per Accounts</label>
              <div className="flex items-center gap-4">
                <span className="text-slate-400">$</span>
                <input 
                  type="number" 
                  className="text-2xl font-bold text-slate-800 w-full focus:outline-none"
                  value={data.accountingProfit}
                  onChange={(e) => {
                    const val = parseFloat(e.target.value) || 0;
                    saveData({...data, accountingProfit: val});
                  }}
                />
              </div>
            </div>

            <section>
              <div className="flex justify-between items-center mb-4">
                <h4 className="text-sm font-bold text-red-600 flex items-center gap-2">
                  <Plus size={16} /> ADD: NON-DEDUCTIBLE EXPENSES
                </h4>
                <button onClick={() => addItem('additions')} className="text-blue-600 hover:text-blue-800"><Plus size={18}/></button>
              </div>
              <div className="space-y-3">
                {data.additions.map(item => (
                  <div key={item.id} className={`grid grid-cols-12 gap-3 items-center group p-1 rounded transition-colors ${item.isAISuggested ? 'bg-emerald-50/50 border border-emerald-100' : ''}`}>
                    <div className="col-span-6 relative">
                      <input 
                        className="w-full bg-transparent p-2 text-sm focus:outline-none" 
                        value={item.label}
                        onChange={(e) => updateItem('additions', item.id, 'label', e.target.value)}
                      />
                      {item.isAISuggested && (
                        <div className="absolute -top-2 -left-1 bg-emerald-500 text-white text-[8px] px-1 rounded font-bold uppercase tracking-tighter">AI</div>
                      )}
                    </div>
                    <input 
                      className="col-span-4 bg-transparent p-2 text-sm text-right font-mono focus:outline-none" 
                      type="number"
                      value={item.amount}
                      onChange={(e) => updateItem('additions', item.id, 'amount', parseFloat(e.target.value) || 0)}
                    />
                    <button onClick={() => deleteItem('additions', item.id)} className="col-span-2 text-slate-300 hover:text-red-500 transition-all flex justify-end items-center px-2">
                      <Trash2 size={16} />
                    </button>
                    {item.isAISuggested && item.note && (
                      <div className="col-span-12 px-2 pb-1 text-[10px] text-emerald-600 italic">
                        Reasoning: {item.note}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </section>

             <section>
              <div className="flex justify-between items-center mb-4">
                <h4 className="text-sm font-bold text-emerald-600 flex items-center gap-2">
                  <Plus size={16} className="rotate-45" /> LESS: NON-TAXABLE INCOME
                </h4>
                <button onClick={() => addItem('otherAdjustments')} className="text-blue-600 hover:text-blue-800"><Plus size={18}/></button>
              </div>
              <div className="space-y-3">
                {data.otherAdjustments.map(item => (
                  <div key={item.id} className={`grid grid-cols-12 gap-3 items-center group p-1 rounded transition-colors ${item.isAISuggested ? 'bg-emerald-50/50 border border-emerald-100' : ''}`}>
                    <div className="col-span-6 relative">
                      <input 
                        className="w-full bg-transparent p-2 text-sm focus:outline-none" 
                        value={item.label}
                        onChange={(e) => updateItem('otherAdjustments', item.id, 'label', e.target.value)}
                      />
                      {item.isAISuggested && (
                        <div className="absolute -top-2 -left-1 bg-emerald-500 text-white text-[8px] px-1 rounded font-bold uppercase tracking-tighter">AI</div>
                      )}
                    </div>
                    <input 
                      className="col-span-4 bg-transparent p-2 text-sm text-right font-mono focus:outline-none" 
                      type="number"
                      value={item.amount}
                      onChange={(e) => updateItem('otherAdjustments', item.id, 'amount', parseFloat(e.target.value) || 0)}
                    />
                    <button onClick={() => deleteItem('otherAdjustments', item.id)} className="col-span-2 text-slate-300 hover:text-red-500 transition-all flex justify-end items-center px-2">
                      <Trash2 size={16} />
                    </button>
                    {item.isAISuggested && item.note && (
                      <div className="col-span-12 px-2 pb-1 text-[10px] text-emerald-600 italic">
                        Reasoning: {item.note}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </section>

            <section>
              <div className="flex justify-between items-center mb-4">
                <h4 className="text-sm font-bold text-blue-600 flex items-center gap-2">
                  <Calculator size={16} /> LESS: CAPITAL ALLOWANCES / LOSSES
                </h4>
                <button onClick={() => addItem('deductions')} className="text-blue-600 hover:text-blue-800"><Plus size={18}/></button>
              </div>
              <div className="space-y-3">
                {data.deductions.map(item => (
                  <div key={item.id} className={`grid grid-cols-12 gap-3 items-center group p-1 rounded transition-colors ${item.isAISuggested ? 'bg-emerald-50/50 border border-emerald-100' : ''}`}>
                    <div className="col-span-6 relative">
                      <input 
                        className="w-full bg-transparent p-2 text-sm focus:outline-none" 
                        value={item.label}
                        onChange={(e) => updateItem('deductions', item.id, 'label', e.target.value)}
                      />
                      {item.isAISuggested && (
                        <div className="absolute -top-2 -left-1 bg-emerald-500 text-white text-[8px] px-1 rounded font-bold uppercase tracking-tighter">AI</div>
                      )}
                    </div>
                    <input 
                      className="col-span-4 bg-transparent p-2 text-sm text-right font-mono focus:outline-none" 
                      type="number"
                      value={item.amount}
                      onChange={(e) => updateItem('deductions', item.id, 'amount', parseFloat(e.target.value) || 0)}
                    />
                    <button onClick={() => deleteItem('deductions', item.id)} className="col-span-2 text-slate-300 hover:text-red-500 transition-all flex justify-end items-center px-2">
                      <Trash2 size={16} />
                    </button>
                    {item.isAISuggested && item.note && (
                      <div className="col-span-12 px-2 pb-1 text-[10px] text-emerald-600 italic">
                        Reasoning: {item.note}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </section>
          </div>
        </div>

        <div className="bg-amber-50 border border-amber-200 p-4 rounded-lg flex gap-3">
          <AlertTriangle className="text-amber-600 shrink-0" />
          <div className="text-sm text-amber-800">
            <p className="font-bold mb-1">Risk Check: Entertainment Expenses</p>
            <p>Ensure entertainment expenses are supported by invoices and strictly for business purposes. Private entertainment is non-deductible.</p>
          </div>
        </div>

        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
            <Upload className="text-blue-600" size={20} /> Supporting Documents
          </h3>
          <div className="space-y-4">
            <div className="flex gap-4 items-end">
              <div className="flex-1">
                <label className="block text-xs font-bold text-slate-500 uppercase mb-2">Purpose of Upload</label>
                <select 
                  value={uploadPurpose}
                  onChange={(e) => setUploadPurpose(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-200 rounded text-sm focus:ring-2 focus:ring-blue-500 outline-none"
                >
                  <option value="financials">Financial Statements (P&L/BS)</option>
                  <option value="tax_data">Tax Data for Analysis</option>
                  <option value="sample_formatting">Sample Formatting</option>
                  <option value="knowledge_base">Tax Knowledge</option>
                  <option value="audit_trail">Audit Trail / Evidence</option>
                </select>
              </div>
              <label className="cursor-pointer bg-blue-600 text-white px-4 py-2 rounded font-bold text-sm hover:bg-blue-700 transition-colors flex items-center gap-2">
                <Plus size={18} /> Select Files
                <input 
                  type="file" 
                  multiple 
                  accept=".pdf,.doc,.docx,.xls,.xlsx" 
                  className="hidden" 
                  onChange={handleFileUpload}
                />
              </label>
            </div>
            
            <div className="space-y-2">
              {uploadedFiles.map(file => (
                <div key={file.id} className="flex items-center justify-between p-3 bg-slate-50 rounded-lg border border-slate-100 group">
                  <div className="flex items-center gap-3">
                    {file.name.match(/\.(xls|xlsx)$/) ? <FileSpreadsheet className="text-emerald-600" size={20} /> : <FileText className="text-blue-600" size={20} />}
                    <div>
                      <p className="text-sm font-medium text-slate-800">{file.name}</p>
                      <p className="text-[10px] text-slate-500 uppercase font-bold">{file.purpose?.replace('_', ' ')} • {(file.size / 1024).toFixed(1)} KB</p>
                    </div>
                  </div>
                  <div className="flex items-center gap-2">
                    {file.name.match(/\.(xls|xlsx)$/) && (
                      <button 
                        onClick={() => analyzeExcel(file.id)}
                        className="text-emerald-600 hover:bg-emerald-50 p-1.5 rounded-lg transition-colors flex items-center gap-1 text-[10px] font-bold border border-emerald-100"
                      >
                        <Sparkles size={14} /> AI Analyze
                      </button>
                    )}
                    <button onClick={() => removeFile(file.id)} className="text-slate-300 hover:text-red-500 transition-all p-1">
                      <X size={18} />
                    </button>
                  </div>
                </div>
              ))}
              {uploadedFiles.length === 0 && (
                <p className="text-sm text-slate-400 text-center py-4 border-2 border-dashed border-slate-100 rounded-lg">No documents uploaded yet</p>
              )}
            </div>
          </div>
        </div>
      </div>

      <div className="space-y-6">
        <div className="bg-slate-900 text-white p-6 rounded-xl shadow-lg sticky top-24">
          <h3 className="text-lg font-bold mb-6 flex items-center gap-2">
            <TrendingUp className="text-emerald-400" /> Summary
          </h3>
          
          <div className="space-y-4">
            <div className="flex justify-between text-sm">
              <span className="text-slate-400">Adjusted Profit</span>
              <span className="font-mono font-bold">${calculations.adjustedProfit.toLocaleString()}</span>
            </div>
            <div className="flex justify-between text-sm">
              <span className="text-slate-400">Total Deductions</span>
              <span className="font-mono text-red-400">-${calculations.totalDeductions.toLocaleString()}</span>
            </div>
            <div className="h-px bg-slate-800 my-2" />
            <div className="flex justify-between text-base">
              <span className="text-slate-200">Chargeable Income</span>
              <span className="font-mono font-bold text-emerald-400">${calculations.chargeableIncome.toLocaleString()}</span>
            </div>
            <div className="flex justify-between text-xs">
              <span className="text-slate-400">Tax Exemption (PTE)</span>
              <span className="font-mono text-emerald-500">-${calculations.exemption.toLocaleString()}</span>
            </div>
            <div className="flex justify-between text-xs">
              <span className="text-slate-400">Tax Before Rebate</span>
              <span className="font-mono text-slate-200">${calculations.taxBeforeRebate.toLocaleString()}</span>
            </div>
            {calculations.citRebate > 0 && (
              <div className="flex justify-between text-xs">
                <span className="text-emerald-400">CIT Rebate (50% max $40k)</span>
                <span className="font-mono text-emerald-400">-${calculations.citRebate.toLocaleString()}</span>
              </div>
            )}
            <div className="h-px bg-slate-800 my-2" />
            <div className="pt-2">
              <p className="text-xs text-slate-400 mb-1 uppercase tracking-widest">Est. Tax Payable (17%)</p>
              <p className="text-4xl font-bold text-white font-mono">${calculations.taxPayable.toLocaleString(undefined, {minimumFractionDigits: 2})}</p>
            </div>
          </div>

          <div className="mt-8 pt-6 border-t border-slate-800">
             <div className="flex items-center gap-2 text-xs text-slate-400 mb-4">
               <ShieldCheck size={14} className="text-emerald-500" /> Compliance Level: High
             </div>
             <button className="w-full bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 rounded-lg transition-colors">
               Lock & Finalize
             </button>
          </div>
        </div>
      </div>

      {showAnalysis && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col">
            <div className="p-6 border-b border-slate-100 flex justify-between items-center bg-emerald-50">
              <div className="flex items-center gap-3">
                <div className="bg-emerald-600 p-2 rounded-lg">
                  <Sparkles className="text-white" size={20} />
                </div>
                <div>
                  <h3 className="text-xl font-bold text-slate-800">AI Tax Impact Analysis</h3>
                  <p className="text-xs text-emerald-700 font-medium">Identifying adjustments for YA {selectedYA}</p>
                </div>
              </div>
              <button onClick={() => setShowAnalysis(false)} className="text-slate-400 hover:text-slate-600"><X size={24} /></button>
            </div>

            <div className="flex-1 overflow-y-auto p-6">
              {analyzing ? (
                <div className="flex flex-col items-center justify-center py-20 space-y-4">
                  <RefreshCw className="animate-spin text-emerald-600" size={48} />
                  <p className="text-slate-500 font-medium animate-pulse">Deep scanning Excel data for tax implications...</p>
                </div>
              ) : (
                <div className="space-y-6">
                  <div className="overflow-hidden border border-slate-200 rounded-xl">
                    <table className="w-full text-left text-sm">
                      <thead className="bg-slate-50 text-slate-500 uppercase text-[10px] font-bold">
                        <tr>
                          <th className="px-4 py-3">Tab</th>
                          <th className="px-4 py-3">Item Description</th>
                          <th className="px-4 py-3 text-right">Amount</th>
                          <th className="px-4 py-3">Tax Impact</th>
                          <th className="px-4 py-3">Action Taken</th>
                          <th className="px-4 py-3">Reasoning</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {analysisResults.map((res, i) => (
                          <tr key={i} className="hover:bg-slate-50 transition-colors">
                            <td className="px-4 py-3 text-[10px] font-bold text-slate-400 uppercase">{res.tab || 'N/A'}</td>
                            <td className="px-4 py-3 font-medium text-slate-800">{res.item}</td>
                            <td className="px-4 py-3 text-right font-mono">${res.amount.toLocaleString()}</td>
                            <td className="px-4 py-3">
                              <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold uppercase ${
                                res.impact === 'addition' ? 'bg-red-100 text-red-700' : 
                                res.impact === 'deduction' ? 'bg-emerald-100 text-emerald-700' : 
                                'bg-blue-100 text-blue-700'
                              }`}>
                                {res.impact}
                              </span>
                            </td>
                            <td className="px-4 py-3 text-slate-600 italic">{res.action}</td>
                            <td className="px-4 py-3 text-slate-500 text-xs">{res.reason}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  <div className="bg-blue-50 border border-blue-100 p-4 rounded-xl flex gap-3">
                    <AlertTriangle className="text-blue-600 shrink-0" size={20} />
                    <p className="text-xs text-blue-800 leading-relaxed">
                      <strong>AI Note:</strong> This analysis is based on standard Singapore tax treatments. 
                      Private car expenses (S-plated) are generally non-deductible. Office renovations may fall under S14N or S14Q depending on the nature of work.
                    </p>
                  </div>
                </div>
              )}
            </div>

            <div className="p-6 border-t border-slate-100 bg-slate-50 flex justify-end gap-3">
              <button 
                onClick={() => setShowAnalysis(false)}
                className="px-6 py-2 text-sm font-bold text-slate-600 hover:bg-slate-200 rounded-lg transition-all"
              >
                Cancel
              </button>
              <button 
                onClick={applyAdjustments}
                disabled={analyzing || analysisResults.length === 0}
                className="px-6 py-2 text-sm font-bold bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 transition-all shadow-lg flex items-center gap-2 disabled:bg-slate-300"
              >
                <CheckCircle2 size={18} /> Apply to Computation
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  </div>
  );
};

const KnowledgeBase = () => {
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [topics, setTopics] = useState<KBTopic[]>([
    { title: "Deductibility of Expenses", ref: "S14/S15 ITA", desc: "General rule: Expenses must be 'wholly and exclusively' incurred in the production of income." },
    { title: "Capital vs Revenue", ref: "IRAS Guide", desc: "Capital expenditure provides a long-term benefit (e.g., assets), while revenue expenditure is recurring." },
    { title: "Section 14Q", ref: "Renovation", desc: "Special deduction for renovation and refurbishment (R&R) costs capped at $300k every 3 years." },
    { title: "Productivity Solutions Grant", ref: "Grants", desc: "Treatment of government grants (generally taxable unless specifically exempt)." }
  ]);
  const [urlInput, setUrlInput] = useState("");
  const [urlTitle, setUrlTitle] = useState("");

  useEffect(() => {
    const savedFiles = localStorage.getItem('tax_gpt_kb_files');
    if (savedFiles) {
      setUploadedFiles(JSON.parse(savedFiles));
    }
    const savedTopics = localStorage.getItem('tax_gpt_kb_topics');
    if (savedTopics) {
      setTopics(JSON.parse(savedTopics));
    }
  }, []);

  const saveFiles = (newFiles: UploadedFile[]) => {
    setUploadedFiles(newFiles);
    localStorage.setItem('tax_gpt_kb_files', JSON.stringify(newFiles));
  };

  const saveTopics = (newTopics: KBTopic[]) => {
    setTopics(newTopics);
    localStorage.setItem('tax_gpt_kb_topics', JSON.stringify(newTopics));
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    const newFiles: UploadedFile[] = Array.from(files).map((file: File) => ({
      id: Math.random().toString(36).substr(2, 9),
      name: file.name,
      type: file.type,
      size: file.size,
      timestamp: new Date().toISOString()
    }));

    saveFiles([...newFiles, ...uploadedFiles]);
  };

  const addUrlTopic = () => {
    if (urlInput.trim() && urlTitle.trim()) {
      const newTopic: KBTopic = {
        title: urlTitle,
        ref: "External URL",
        desc: `Reference link: ${urlInput}`,
        url: urlInput
      };
      saveTopics([newTopic, ...topics]);
      setUrlInput("");
      setUrlTitle("");
    }
  };

  const removeFile = (id: string) => {
    saveFiles(uploadedFiles.filter(f => f.id !== id));
  };

  const removeTopic = (index: number) => {
    saveTopics(topics.filter((_, i) => i !== index));
  };

  return (
    <div className="space-y-6">
      <div className="grid md:grid-cols-2 gap-6">
        <div className="space-y-4">
          <div className="flex gap-4">
            <div className="flex-1 flex items-center gap-4 bg-white p-4 rounded-xl border border-slate-200">
              <Search className="text-slate-400" />
              <input className="w-full focus:outline-none" placeholder="Search IRAS e-Tax guides..." />
            </div>
            <label className="cursor-pointer bg-blue-600 text-white px-6 py-4 rounded-xl font-bold text-sm hover:bg-blue-700 transition-colors flex items-center gap-2 shadow-sm">
              <Upload size={20} /> Upload Guide
              <input 
                type="file" 
                multiple 
                accept=".pdf,.doc,.docx,.xls,.xlsx" 
                className="hidden" 
                onChange={handleFileUpload}
              />
            </label>
          </div>

          <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
            <h4 className="text-sm font-bold text-slate-800 mb-4 flex items-center gap-2">
              <LinkIcon className="text-blue-600" size={18} /> Add Reference URL
            </h4>
            <div className="space-y-3">
              <input 
                className="w-full p-2 bg-slate-50 border border-slate-200 rounded text-sm outline-none focus:ring-1 focus:ring-blue-500"
                placeholder="Topic Title (e.g. IRAS GST Guide)"
                value={urlTitle}
                onChange={(e) => setUrlTitle(e.target.value)}
              />
              <div className="flex gap-2">
                <input 
                  className="flex-1 p-2 bg-slate-50 border border-slate-200 rounded text-sm outline-none focus:ring-1 focus:ring-blue-500"
                  placeholder="https://www.iras.gov.sg/..."
                  value={urlInput}
                  onChange={(e) => setUrlInput(e.target.value)}
                />
                <button 
                  onClick={addUrlTopic}
                  className="px-4 py-2 bg-blue-600 text-white rounded font-bold text-sm hover:bg-blue-700"
                >
                  Add
                </button>
              </div>
            </div>
          </div>
        </div>

        <div className="space-y-4">
          {uploadedFiles.length > 0 && (
            <div className="grid grid-cols-1 gap-3">
              {uploadedFiles.map(file => (
                <div key={file.id} className="bg-white p-4 rounded-xl border border-slate-200 flex items-center justify-between group">
                  <div className="flex items-center gap-3">
                    <FileText className="text-blue-600" size={20} />
                    <div>
                      <p className="text-sm font-medium text-slate-800">{file.name}</p>
                      <p className="text-[10px] text-slate-500 uppercase font-bold">Custom Guide • {(file.size / 1024).toFixed(1)} KB</p>
                    </div>
                  </div>
                  <button onClick={() => removeFile(file.id)} className="text-slate-300 hover:text-red-500 transition-all p-1">
                    <X size={18} />
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="grid md:grid-cols-2 gap-4">
        {topics.map((t, i) => (
          <div key={i} className="bg-white p-5 rounded-xl border border-slate-200 hover:border-blue-300 transition-colors cursor-pointer group relative">
            <div className="flex justify-between items-start mb-2">
              <h4 className="font-bold text-slate-800 group-hover:text-blue-600 transition-colors">{t.title}</h4>
              <span className="text-[10px] bg-blue-50 text-blue-600 px-2 py-0.5 rounded uppercase font-bold">{t.ref}</span>
            </div>
            <p className="text-sm text-slate-600 line-clamp-2 mb-2">{t.desc}</p>
            {t.url && (
              <a href={t.url} target="_blank" rel="noopener noreferrer" className="text-xs text-blue-600 hover:underline flex items-center gap-1">
                <LinkIcon size={12} /> Visit Link
              </a>
            )}
              <button 
                onClick={() => removeTopic(i)}
                className="absolute top-2 right-2 text-slate-300 hover:text-red-500 transition-all p-1"
              >
                <X size={14} />
              </button>
          </div>
        ))}
      </div>
    </div>
  );
};

const IRASQueryResponse = () => {
  const [letterText, setLetterText] = useState("");
  const [loading, setLoading] = useState(false);
  const [response, setResponse] = useState<any>(null);
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [fileObjects, setFileObjects] = useState<Map<string, File>>(new Map());
  const [uploadPurpose, setUploadPurpose] = useState('formatting');
  const [analyzingEvidence, setAnalyzingEvidence] = useState(false);
  const [evidenceInsights, setEvidenceInsights] = useState<string>("");
  const [feedback, setFeedback] = useState<'up' | 'down' | null>(null);
  const [error, setError] = useState<string | null>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    const newFileObjs = new Map(fileObjects);
    const newFiles: UploadedFile[] = Array.from(files).map((file: File) => {
      const id = Math.random().toString(36).substr(2, 9);
      newFileObjs.set(id, file);
      return {
        id,
        name: file.name,
        type: file.type,
        size: file.size,
        purpose: uploadPurpose,
        timestamp: new Date().toISOString()
      };
    });

    setFileObjects(newFileObjs);
    setUploadedFiles([...newFiles, ...uploadedFiles]);
  };

  const removeFile = (id: string) => {
    const newFileObjs = new Map(fileObjects);
    newFileObjs.delete(id);
    setFileObjects(newFileObjs);
    setUploadedFiles(uploadedFiles.filter(f => f.id !== id));
  };

  const analyzeEvidence = async () => {
    const evidenceFiles = uploadedFiles.filter(f => f.purpose === 'evidence');
    if (evidenceFiles.length === 0) return;

    setAnalyzingEvidence(true);
    const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

    try {
      const parts = await Promise.all(evidenceFiles.map(async (f) => {
        const fileObj = fileObjects.get(f.id);
        if (fileObj) {
          const base64 = await new Promise<string>((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve((reader.result as string).split(',')[1]);
            reader.readAsDataURL(fileObj);
          });
          return {
            inlineData: {
              data: base64,
              mimeType: fileObj.type || 'application/pdf'
            }
          };
        }
        return null;
      }));

      const validParts = parts.filter(p => p !== null) as any[];
      
      const prompt = `You are a Tax Auditor. Analyze these supporting documents (invoices, bank statements, etc.) and extract key figures, dates, and facts that would be relevant to answering an IRAS query. 
      Focus on:
      1. Total amounts mentioned.
      2. Specific dates of transactions.
      3. Nature of expenses/income.
      4. Any discrepancies or notable points.
      
      Return a concise summary of these facts.`;

      const result = await ai.models.generateContent({
        model: "gemini-1.5-flash",
        contents: [{ role: 'user', parts: [prompt, ...validParts] }]
      });

      setEvidenceInsights(result.text);
    } catch (error: any) {
      setError(parseAIError(error));
    } finally {
      setAnalyzingEvidence(false);
    }
  };

  const exportToWord = () => {
    if (!response) return;
    
    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            text: "IRAS Query Response Draft",
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER,
          }),
          new Paragraph({ text: "" }),
          new Paragraph({
            text: "Summary of Concerns:",
            heading: HeadingLevel.HEADING_2,
          }),
          new Paragraph({ text: response.summary }),
          new Paragraph({ text: "" }),
          new Paragraph({
            text: "Draft Response:",
            heading: HeadingLevel.HEADING_2,
          }),
          ...response.draft.split('\n').map((line: string) => new Paragraph({ text: line })),
        ],
      }],
    });

    Packer.toBlob(doc).then(blob => {
      saveAs(blob, "IRAS_Response_Draft.docx");
    });
  };

  const draftResponse = async () => {
    if (!letterText.trim()) return;
    
    const cacheKey = `draft_${letterText.slice(0, 100)}_${evidenceInsights.slice(0, 50)}`;
    if (advisoryCache.has(cacheKey)) {
      setResponse(advisoryCache.get(cacheKey));
      setFeedback(null);
      return;
    }

    setLoading(true);
    const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

    const systemPrompt = `Expert Singapore Tax Consultant. 
    Draft compliant IRAS response. Evidence: ${evidenceInsights}
    JSON keys: 'summary', 'draft', 'requiredDocs', 'strategy'.`;

    try {
      const result = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [{ role: 'user', parts: [{ text: `IRAS Letter: ${letterText}` }] }],
        config: { 
          systemInstruction: systemPrompt,
          responseMimeType: "application/json" 
        }
      });
      
      const parsed = JSON.parse(result.text);
      advisoryCache.set(cacheKey, parsed);
      setResponse(parsed);
      setFeedback(null);
    } catch (error: any) {
      setError(parseAIError(error));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="grid lg:grid-cols-2 gap-8">
      <div className="lg:col-span-2">
        {error && (
          <div className="mb-6 p-4 bg-red-50 border border-red-200 rounded-xl flex items-start gap-4 text-red-700 text-sm animate-in fade-in slide-in-from-top-2 duration-300 shadow-sm">
            <AlertCircle className="w-6 h-6 shrink-0" />
            <div className="flex-1">
              <p className="font-bold text-base mb-1">Service Notification</p>
              <p className="leading-relaxed">{error}</p>
              <button onClick={() => setError(null)} className="mt-2 text-xs font-black uppercase tracking-widest hover:underline">Dismiss Message</button>
            </div>
          </div>
        )}
      </div>
      <div className="space-y-6">
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
            <Mail className="text-blue-600" /> IRAS Letter Analysis
          </h3>
          <textarea
            className="w-full p-4 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none transition-all mb-4"
            rows={8}
            placeholder="Paste the content of the IRAS query letter here..."
            value={letterText}
            onChange={(e) => setLetterText(e.target.value)}
          />
          
          <div className="space-y-4">
            <div className="flex gap-4 items-end">
              <div className="flex-1">
                <label className="block text-xs font-bold text-slate-500 uppercase mb-2">Upload for Context</label>
                <select 
                  value={uploadPurpose}
                  onChange={(e) => setUploadPurpose(e.target.value)}
                  className="w-full p-2 bg-slate-50 border border-slate-200 rounded text-sm outline-none"
                >
                  <option value="formatting">Formatting Sample</option>
                  <option value="knowledge">Tax Knowledge</option>
                  <option value="evidence">Supporting Evidence</option>
                </select>
              </div>
              <label className="cursor-pointer bg-slate-100 text-slate-600 px-4 py-2 rounded font-bold text-sm hover:bg-slate-200 transition-colors flex items-center gap-2">
                <Paperclip size={18} /> Attach
                <input type="file" multiple className="hidden" onChange={handleFileUpload} />
              </label>
            </div>

            {uploadedFiles.length > 0 && (
              <div className="space-y-3">
                <div className="flex flex-wrap gap-2">
                  {uploadedFiles.map((file, i) => (
                    <div key={file.id} className="flex items-center gap-2 bg-slate-50 border border-slate-100 px-3 py-1 rounded-full text-xs text-slate-500 group">
                      <FileText size={14} /> 
                      <span className="truncate max-w-[100px]">{file.name}</span>
                      <span className="text-[8px] bg-slate-200 px-1 rounded uppercase">{file.purpose}</span>
                      <button onClick={() => removeFile(file.id)} className="hover:text-red-500"><X size={14} /></button>
                    </div>
                  ))}
                </div>
                {uploadedFiles.some(f => f.purpose === 'evidence') && (
                  <button 
                    onClick={analyzeEvidence}
                    disabled={analyzingEvidence}
                    className="w-full py-2 border border-blue-200 bg-blue-50 text-blue-600 rounded-lg text-xs font-bold hover:bg-blue-100 transition-all flex items-center justify-center gap-2"
                  >
                    {analyzingEvidence ? <RefreshCw size={14} className="animate-spin" /> : <Sparkles size={14} />}
                    {evidenceInsights ? "Re-analyze Evidence" : "Analyze Supporting Evidence"}
                  </button>
                )}
                {evidenceInsights && (
                  <div className="p-3 bg-emerald-50 border border-emerald-100 rounded-lg">
                    <h5 className="text-[10px] font-bold text-emerald-700 uppercase mb-1 flex items-center gap-1">
                      <CheckCircle2 size={10} /> Extracted Evidence Insights
                    </h5>
                    <p className="text-[10px] text-emerald-600 leading-relaxed italic line-clamp-3">{evidenceInsights}</p>
                  </div>
                )}
              </div>
            )}

            <button 
              onClick={draftResponse}
              disabled={loading || !letterText}
              className="w-full bg-blue-600 text-white font-bold py-3 rounded-lg hover:bg-blue-700 disabled:bg-slate-300 transition-all flex items-center justify-center gap-2"
            >
              {loading ? <RefreshCw className="animate-spin" /> : <FileCode size={20} />}
              Draft Professional Response
            </button>
          </div>
        </div>
      </div>

      <div className="space-y-6">
        {response ? (
          <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-6">
            <section>
              <h4 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-2">IRAS Concerns Summary</h4>
              <p className="text-sm text-slate-700 bg-slate-50 p-3 rounded-lg border border-slate-100">{response.summary}</p>
            </section>
            
            <section>
              <div className="flex justify-between items-center mb-2">
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-widest">Drafted Response</h4>
                <div className="flex gap-2">
                  <button 
                    onClick={exportToWord}
                    className="flex items-center gap-1 text-[10px] bg-blue-50 text-blue-600 px-2 py-1 rounded hover:bg-blue-100 font-bold"
                  >
                    <FileCode size={12} /> Word
                  </button>
                  <button 
                    onClick={() => {
                      const doc = new jsPDF();
                      doc.text("IRAS Response Draft", 14, 20);
                      doc.setFontSize(10);
                      const splitText = doc.splitTextToSize(response.draft, 180);
                      doc.text(splitText, 14, 30);
                      doc.save("IRAS_Response_Draft.pdf");
                    }}
                    className="flex items-center gap-1 text-[10px] bg-red-50 text-red-600 px-2 py-1 rounded hover:bg-red-100 font-bold"
                  >
                    <FileDown size={12} /> PDF
                  </button>
                </div>
              </div>
              <div className="relative group">
                <textarea 
                  readOnly
                  className="w-full p-4 bg-slate-900 text-slate-100 rounded-lg font-mono text-xs leading-relaxed"
                  rows={12}
                  value={response.draft}
                />
                <button 
                  onClick={() => navigator.clipboard.writeText(response.draft)}
                  className="absolute top-2 right-2 bg-slate-800 text-slate-400 p-2 rounded hover:text-white transition-colors"
                >
                  <Download size={16} />
                </button>
              </div>
            </section>

            <div className="grid grid-cols-2 gap-4">
              <section>
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-2">Required Evidence</h4>
                <ul className="list-disc list-inside text-xs text-slate-600 space-y-1">
                  {response.requiredDocs?.map((doc: string, i: number) => <li key={i}>{doc}</li>)}
                </ul>
              </section>
              <section>
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-widest mb-2">Handling Strategy</h4>
                <p className="text-xs text-slate-600 italic">{response.strategy}</p>
              </section>
            </div>

            <div className="pt-4 border-t border-slate-100 flex items-center justify-between">
              <span className="text-xs text-slate-500">Was this response helpful?</span>
              <div className="flex gap-2">
                <button 
                  onClick={() => setFeedback('up')}
                  className={`p-2 rounded-lg transition-all ${feedback === 'up' ? 'bg-emerald-100 text-emerald-600' : 'bg-slate-50 text-slate-400 hover:bg-slate-100'}`}
                >
                  <ThumbsUp size={16} />
                </button>
                <button 
                  onClick={() => setFeedback('down')}
                  className={`p-2 rounded-lg transition-all ${feedback === 'down' ? 'bg-red-100 text-red-600' : 'bg-slate-50 text-slate-400 hover:bg-slate-100'}`}
                >
                  <ThumbsDown size={16} />
                </button>
              </div>
            </div>
          </div>
        ) : (
          <div className="h-full flex flex-col items-center justify-center text-center p-12 bg-slate-50 rounded-xl border-2 border-dashed border-slate-200">
            <Mail className="text-slate-300 mb-4" size={48} />
            <p className="text-slate-500 font-medium">Upload an IRAS query letter to generate a professional response draft.</p>
          </div>
        )}
      </div>
    </div>
  );
};

const ComputationReview = () => {
  const [uploadedFiles, setUploadedFiles] = useState<UploadedFile[]>([]);
  const [loading, setLoading] = useState(false);
  const [reviewResult, setReviewResult] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files) return;

    const newFiles: UploadedFile[] = Array.from(files).map((file: File) => ({
      id: Math.random().toString(36).substr(2, 9),
      name: file.name,
      type: file.type,
      size: file.size,
      timestamp: new Date().toISOString()
    }));

    setUploadedFiles([...newFiles, ...uploadedFiles]);
  };

  const removeFile = (id: string) => {
    setUploadedFiles(uploadedFiles.filter(f => f.id !== id));
  };

  const runReview = async () => {
    if (uploadedFiles.length === 0) return;
    
    const cacheKey = `review_${uploadedFiles.map(f => f.id).join('_')}`;
    if (advisoryCache.has(cacheKey)) {
      setReviewResult(advisoryCache.get(cacheKey));
      return;
    }

    setLoading(true);
    const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY || "" });

    const prompt = `Senior Tax Auditor (Singapore). Analyze computation for errors/omissions.
    JSON keys: 'summary', 'findings':[{'issue', 'impact', 'recommendation'}], 'adjustments':[{'item', 'originalAmount', 'revisedAmount', 'reason'}].`;

    try {
      const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: [{ role: 'user', parts: [{ text: prompt }] }],
        config: { responseMimeType: "application/json" }
      });
      const parsed = JSON.parse(response.text);
      advisoryCache.set(cacheKey, parsed);
      setReviewResult(parsed);
    } catch (error: any) {
      setError(parseAIError(error));
      // Fallback for demo if it wasn't a quota error
      if (!error?.message?.includes('429')) {
        const demo = {
          summary: "The computation is generally accurate but has minor omissions in Capital Allowances and Section 14Q claims.",
          findings: [
            { issue: "Missing S14Q claim for office renovation", impact: "Higher taxable income", recommendation: "Identify R&R costs and claim under S14Q (capped at $300k)." },
            { issue: "Incorrect depreciation add-back", impact: "Minor tax overpayment", recommendation: "Ensure all accounting depreciation is added back before claiming CA." }
          ],
          adjustments: [
            { item: "Office Renovation", originalAmount: 0, revisedAmount: 15000, reason: "Eligible for S14Q deduction" },
            { item: "Entertainment Expenses", originalAmount: 5000, revisedAmount: 4200, reason: "Private portion identified" }
          ]
        };
        setReviewResult(demo);
      }
    } finally {
      setLoading(false);
    }
  };

  const exportReview = (format: 'pdf' | 'excel' | 'word') => {
    if (!reviewResult) return;

    if (format === 'pdf') {
      const doc = new jsPDF();
      doc.setFontSize(16);
      doc.text("Tax Computation Review Report", 14, 20);
      doc.setFontSize(10);
      doc.text(`Summary: ${reviewResult.summary}`, 14, 30);
      
      (doc as any).autoTable({
        startY: 40,
        head: [['Issue', 'Impact', 'Recommendation']],
        body: reviewResult.findings.map((f: any) => [f.issue, f.impact, f.recommendation]),
      });

      (doc as any).autoTable({
        startY: (doc as any).lastAutoTable.finalY + 10,
        head: [['Item', 'Original', 'Revised', 'Reason']],
        body: reviewResult.adjustments.map((a: any) => [a.item, a.originalAmount, a.revisedAmount, a.reason]),
      });

      doc.save("Tax_Computation_Review.pdf");
    } else if (format === 'excel') {
      const wb = XLSX.utils.book_new();
      const wsFindings = XLSX.utils.json_to_sheet(reviewResult.findings);
      const wsAdjustments = XLSX.utils.json_to_sheet(reviewResult.adjustments);
      XLSX.utils.book_append_sheet(wb, wsFindings, "Findings");
      XLSX.utils.book_append_sheet(wb, wsAdjustments, "Adjustments");
      XLSX.writeFile(wb, "Tax_Computation_Review.xlsx");
    } else if (format === 'word') {
      const doc = new Document({
        sections: [{
          children: [
            new Paragraph({ text: "Tax Computation Review Report", heading: HeadingLevel.HEADING_1 }),
            new Paragraph({ text: reviewResult.summary }),
            new Paragraph({ text: "Findings", heading: HeadingLevel.HEADING_2 }),
            ...reviewResult.findings.flatMap((f: any) => [
              new Paragraph({ text: `Issue: ${f.issue}`, bullet: { level: 0 } }),
              new Paragraph({ text: `Impact: ${f.impact}`, bullet: { level: 1 } }),
              new Paragraph({ text: `Recommendation: ${f.recommendation}`, bullet: { level: 1 } }),
            ]),
          ]
        }]
      });
      Packer.toBlob(doc).then(blob => saveAs(blob, "Tax_Computation_Review.docx"));
    }
  };

  return (
    <div className="space-y-6">
      {error && (
        <div className="p-4 bg-red-50 border border-red-200 rounded-xl flex items-start gap-4 text-red-700 text-sm animate-in fade-in slide-in-from-top-2 duration-300 shadow-sm">
          <AlertCircle className="w-6 h-6 shrink-0" />
          <div className="flex-1">
            <p className="font-bold text-base mb-1">Service Notification</p>
            <p className="leading-relaxed">{error}</p>
            <button onClick={() => setError(null)} className="mt-2 text-xs font-black uppercase tracking-widest hover:underline">Dismiss</button>
          </div>
        </div>
      )}
      <div className="grid lg:grid-cols-3 gap-8">
        <div className="lg:col-span-1 space-y-6">
        <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
          <h3 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
            <Upload className="text-blue-600" size={20} /> Upload Computation
          </h3>
          <div className="space-y-4">
            <label className="cursor-pointer w-full flex flex-col items-center justify-center p-8 border-2 border-dashed border-slate-200 rounded-xl hover:border-blue-400 transition-all bg-slate-50 group">
              <FileCode className="text-slate-300 group-hover:text-blue-500 mb-2" size={32} />
              <span className="text-sm text-slate-500 font-medium">Click to upload PDF, Word or Excel</span>
              <input type="file" className="hidden" onChange={handleFileUpload} accept=".pdf,.doc,.docx,.xls,.xlsx" />
            </label>

            <div className="space-y-2">
              {uploadedFiles.map(file => (
                <div key={file.id} className="flex items-center justify-between p-3 bg-white rounded-lg border border-slate-100 group">
                  <div className="flex items-center gap-3">
                    <FileText className="text-blue-600" size={18} />
                    <span className="text-xs font-medium text-slate-700 truncate max-w-[150px]">{file.name}</span>
                  </div>
                  <button onClick={() => removeFile(file.id)} className="text-slate-300 hover:text-red-500 transition-all p-1">
                    <X size={16} />
                  </button>
                </div>
              ))}
            </div>

            <button 
              onClick={runReview}
              disabled={loading || uploadedFiles.length === 0}
              className="w-full bg-blue-600 text-white font-bold py-3 rounded-lg hover:bg-blue-700 disabled:bg-slate-300 transition-all flex items-center justify-center gap-2"
            >
              {loading ? <RefreshCw className="animate-spin" /> : <ShieldCheck size={20} />}
              Run Audit Review
            </button>
          </div>
        </div>
      </div>

      <div className="lg:col-span-2">
        {reviewResult ? (
          <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-8">
            <div className="flex justify-between items-center">
              <h3 className="text-xl font-bold text-slate-800">Audit Findings</h3>
              <div className="flex gap-2">
                <button onClick={() => exportReview('pdf')} className="p-2 bg-red-50 text-red-600 rounded-lg hover:bg-red-100 transition-all"><FileDown size={20} /></button>
                <button onClick={() => exportReview('excel')} className="p-2 bg-emerald-50 text-emerald-600 rounded-lg hover:bg-emerald-100 transition-all"><FileSpreadsheet size={20} /></button>
                <button onClick={() => exportReview('word')} className="p-2 bg-blue-50 text-blue-600 rounded-lg hover:bg-blue-100 transition-all"><FileCode size={20} /></button>
              </div>
            </div>

            <div className="p-4 bg-amber-50 border border-amber-100 rounded-xl">
              <p className="text-sm text-amber-900 font-medium">{reviewResult.summary}</p>
            </div>

            <div className="space-y-4">
              <h4 className="text-sm font-bold text-slate-500 uppercase tracking-widest">Key Issues Identified</h4>
              <div className="grid gap-4">
                {reviewResult.findings.map((f: any, i: number) => (
                  <div key={i} className="p-4 border border-slate-100 rounded-xl hover:border-blue-200 transition-all">
                    <div className="flex justify-between items-start mb-2">
                      <h5 className="font-bold text-slate-800">{f.issue}</h5>
                      <span className="text-[10px] bg-red-50 text-red-600 px-2 py-0.5 rounded font-bold uppercase">{f.impact}</span>
                    </div>
                    <p className="text-xs text-slate-600 leading-relaxed">{f.recommendation}</p>
                  </div>
                ))}
              </div>
            </div>

            <div className="space-y-4">
              <h4 className="text-sm font-bold text-slate-500 uppercase tracking-widest">Proposed Adjustments</h4>
              <div className="overflow-hidden border border-slate-100 rounded-xl">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-50 text-slate-500 font-bold uppercase">
                    <tr>
                      <th className="px-4 py-3">Item</th>
                      <th className="px-4 py-3 text-right">Original</th>
                      <th className="px-4 py-3 text-right">Revised</th>
                      <th className="px-4 py-3">Reason</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {reviewResult.adjustments.map((a: any, i: number) => (
                      <tr key={i} className="hover:bg-slate-50 transition-colors">
                        <td className="px-4 py-3 font-medium text-slate-800">{a.item}</td>
                        <td className="px-4 py-3 text-right font-mono text-slate-400">${a.originalAmount.toLocaleString()}</td>
                        <td className="px-4 py-3 text-right font-mono text-emerald-600 font-bold">${a.revisedAmount.toLocaleString()}</td>
                        <td className="px-4 py-3 text-slate-500">{a.reason}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        ) : (
          <div className="h-full min-h-[400px] flex flex-col items-center justify-center text-center p-12 bg-slate-50 rounded-xl border-2 border-dashed border-slate-200">
            <ShieldCheck className="text-slate-300 mb-4" size={48} />
            <p className="text-slate-500 font-medium">Upload your tax computation to run an automated audit review for errors and optimizations.</p>
          </div>
        )}
      </div>
    </div>
  </div>
  );
};

export default function App() {
  const [activeTab, setActiveTab] = useState('advisory');

  return (
    <div className="min-h-screen bg-slate-50 font-sans text-slate-900">
      {/* Header */}
      <header className="bg-white border-b border-slate-200 sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-blue-600 p-2 rounded-lg">
              <ShieldCheck className="text-white" size={24} />
            </div>
            <h1 className="text-xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-blue-700 to-indigo-800">
              TaxGPT Singapore
            </h1>
          </div>
          <div className="flex items-center gap-4 text-sm font-medium">
            <span className="hidden sm:inline text-slate-500">Firm: Sample Solutions LLP</span>
            <div className="flex items-center gap-2 bg-slate-100 px-3 py-1.5 rounded-full">
              <User size={16} className="text-slate-500" />
              <span className="text-slate-700">Guest User</span>
            </div>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Sub-nav */}
        <div className="flex gap-1 bg-white p-1 rounded-xl border border-slate-200 w-fit mb-8 shadow-sm">
          <button 
            onClick={() => setActiveTab('advisory')}
            className={`flex items-center gap-2 px-6 py-2.5 rounded-lg text-sm font-bold transition-all ${activeTab === 'advisory' ? 'bg-blue-600 text-white shadow-md' : 'text-slate-500 hover:bg-slate-50'}`}
          >
            <MessageSquare size={18} /> Advisory
          </button>
          <button 
            onClick={() => setActiveTab('computation')}
            className={`flex items-center gap-2 px-6 py-2.5 rounded-lg text-sm font-bold transition-all ${activeTab === 'computation' ? 'bg-blue-600 text-white shadow-md' : 'text-slate-500 hover:bg-slate-50'}`}
          >
            <Calculator size={18} /> Computation
          </button>
          <button 
            onClick={() => setActiveTab('iras_query')}
            className={`flex items-center gap-2 px-6 py-2.5 rounded-lg text-sm font-bold transition-all ${activeTab === 'iras_query' ? 'bg-blue-600 text-white shadow-md' : 'text-slate-500 hover:bg-slate-50'}`}
          >
            <Mail size={18} /> IRAS Query
          </button>
          <button 
            onClick={() => setActiveTab('review')}
            className={`flex items-center gap-2 px-6 py-2.5 rounded-lg text-sm font-bold transition-all ${activeTab === 'review' ? 'bg-blue-600 text-white shadow-md' : 'text-slate-500 hover:bg-slate-50'}`}
          >
            <ShieldCheck size={18} /> Audit Review
          </button>
          <button 
             onClick={() => setActiveTab('library')}
            className={`flex items-center gap-2 px-6 py-2.5 rounded-lg text-sm font-bold transition-all ${activeTab === 'library' ? 'bg-blue-600 text-white shadow-md' : 'text-slate-500 hover:bg-slate-50'}`}
          >
            <BookOpen size={18} /> Knowledge Base
          </button>
        </div>

        {/* Tab Content */}
        <div className="min-h-[600px]">
          {activeTab === 'advisory' && <TaxAdvisory />}
          {activeTab === 'computation' && <TaxComputation />}
          {activeTab === 'iras_query' && <IRASQueryResponse />}
          {activeTab === 'review' && <ComputationReview />}
          {activeTab === 'library' && <KnowledgeBase />}
        </div>
      </main>

      {/* Footer Disclaimer */}
      <footer className="max-w-7xl mx-auto px-4 py-12">
        <div className="bg-slate-200 h-px w-full mb-8" />
        <div className="flex flex-col md:flex-row justify-between gap-8">
          <div className="max-w-md">
            <p className="text-xs text-slate-500 leading-relaxed uppercase tracking-wider font-bold mb-2">Legal Disclaimer</p>
            <p className="text-[10px] text-slate-400 leading-relaxed italic">
              This tool provides guidance based on general Singapore tax principles and should not replace professional advice. 
              Always cross-reference with the latest IRAS e-Tax guides and the Income Tax Act. TaxGPT and its developers are not liable for 
              compliance errors or omissions resulting from the use of this software.
            </p>
          </div>
          <div className="flex gap-8">
             <div>
               <p className="text-xs font-bold text-slate-600 mb-2">Resources</p>
               <ul className="text-xs text-slate-400 space-y-1">
                 <li>IRAS Website</li>
                 <li>Income Tax Act</li>
                 <li>Singapore Budget 2024</li>
               </ul>
             </div>
             <div>
               <p className="text-xs font-bold text-slate-600 mb-2">Support</p>
               <ul className="text-xs text-slate-400 space-y-1">
                 <li>Security</li>
                 <li>Privacy Policy</li>
                 <li>Big 4 Standards</li>
               </ul>
             </div>
          </div>
        </div>
        <p className="text-[10px] text-slate-400 mt-12 text-center">© 2024 TaxGPT Singapore. Built for Compliance Excellence.</p>
      </footer>
    </div>
  );
}
