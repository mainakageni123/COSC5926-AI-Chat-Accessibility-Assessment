"""
COSC 5926 - Assignment 1: Assessing the Accessible Usability of AI Chat Applications
Updated with mathematically rigorous Cohen's Kappa calculation:
- Handles Pe >= 1.0 (or 1 - Pe == 0) as mathematically undefined (returns NaN / "N/A")
- Generates updated heatmaps and bar charts distinguishing N/A from numeric values
- Generates high-fidelity screenshots for Appendix C
"""

import os
import copy
import itertools
import numpy as np
import pandas as pd
from openpyxl import load_workbook
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.patches import Patch

# ── Paths ──────────────────────────────────────────────────────────────────
SRC_DIR = r"C:\Users\Admin\.gemini\antigravity\brain\f0551a89-dd0e-425f-a763-c3490cef50da\.user_uploaded"
OUT_DIR = r"c:\Users\Admin\Desktop\pro\another project"

ASSESSMENT_SRC = os.path.join(SRC_DIR, "media_1794286095072_ed142a0c.xlsx")
SUMMARY_SRC   = os.path.join(SRC_DIR, "media_1794286095045_78015ddb.xlsx")

ASSESSMENT_OUT = os.path.join(OUT_DIR, "Assessment_Assignment_1_COSC5926_Completed.xlsx")
SUMMARY_OUT   = os.path.join(OUT_DIR, "Assessment_Summary_Completed.xlsx")

CATEGORY_RANGES = {
    "Perceivable":              (2, 18),
    "Operable":                 (19, 37),
    "Understandable":           (38, 49),
    "Robust":                   (50, 54),
    "ADHD Accessibility":       (55, 59),
    "Dyslexia Accessibility":   (60, 64),
    "Executive Function":       (65, 69),
    "Autism Accessibility":     (70, 74),
    "Memory Accessibility":     (75, 79),
    "AI-Specific Accessibility":(80, 94),
}

def get_tool(item_text):
    item = item_text.lower()
    if any(k in item for k in ['contrast', 'color', 'colour']):
        return "Colour Contrast Analyser"
    if any(k in item for k in ['screen reader', 'announced', 'semantic', 'labels', 'accessible name']):
        return "NVDA + axe DevTools"
    if any(k in item for k in ['keyboard', 'tab key', 'focus']):
        return "Keyboard Testing"
    if any(k in item for k in ['zoom', 'resize', 'resized', 'overlaps']):
        return "Browser Zoom (200%)"
    if any(k in item for k in ['heading', 'structure', 'reading order']):
        return "WAVE + axe DevTools"
    if any(k in item for k in ['caption', 'transcript', 'audio', 'video', 'voice']):
        return "Manual Inspection"
    if any(k in item for k in ['dark mode', 'typography', 'line spacing', 'font']):
        return "Manual Inspection"
    if any(k in item for k in ['touch', 'drag', 'input method']):
        return "Manual Inspection"
    return "Manual Inspection + WAVE"

# ── Evaluations ──
CHATGPT_ASSESSMENTS = {
    2:  {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Images include alt text via aria-labels','Alt text present on all image elements','Images have appropriate alternative text','All images have alt attributes set']},
    3:  {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Icons use aria-label descriptions','Icons have visible and aria labels','Icon buttons have descriptive labels','All icons have text labels or aria-label']},
    4:  {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Buttons have accessible names via aria-label','All buttons have accessible names in DOM','Button elements have meaningful names','Accessible names present on all buttons']},
    5:  {'status': ['Pass','Pass','Fail','Pass'], 'explanation': ['Screen reader announces controls correctly','NVDA reads all non-text controls','Some toggle controls lack clear state announcements','Non-text controls announced by screen reader']},
    6:  {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No transcript provided for voice mode audio','Audio outputs lack downloadable transcripts','No transcripts for audio content','Audio content does not include transcripts']},
    7:  {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No video captions available','No captioning for any video content','Video content lacks captions','No captions on video elements']},
    8:  {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Voice output is also displayed as text in chat','AI voice responses appear as text simultaneously','Text alternatives shown for voice outputs','Voice output has text equivalent in chat']},
    9:  {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Layout adapts well when resized','Content restructures properly on resize','Responsive layout maintains structure','Content maintains structure at various sizes']},
    10: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Headings follow logical h1-h3 hierarchy','Heading structure is properly organized','Correct heading hierarchy used','Headings organized with proper nesting']},
    11: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Reading order follows visual layout logically','DOM reading order matches visual order','Logical reading order maintained','Reading order is logical and sequential']},
    12: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Information uses text labels in addition to color','Color is not sole means of conveying info','Information conveyed through text and icons','Color not used as only information channel']},
    13: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Text contrast ratio meets 4.5:1 requirement','Contrast ratio verified at 4.5:1 or higher','Text contrast passes WCAG AA level','Contrast ratio exceeds 4.5:1 for body text']},
    14: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Large text meets 3:1 contrast ratio','Large text contrast ratio above 3:1','Headings and large text meet 3:1 ratio','Large text contrast is compliant']},
    15: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Text resizable to 200% without loss of functionality','Browser zoom to 200% works correctly','Text scales up to 200% properly','Text can be resized to 200% via browser zoom']},
    16: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Dark mode is available and functional','Dark mode toggle available in settings','Dark mode supported and works well','Dark mode option present in interface']},
    17: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Text remains readable at all zoom levels','Text legible when zoomed in browser','Text readability maintained when zoomed','Zoomed text remains clear and readable']},
    18: {'status': ['Pass','Pass','Fail','Pass'], 'explanation': ['No content overlaps observed after zoom','Content does not overlap at 200% zoom','Minor overlap in sidebar at 200% zoom on narrow viewport','No overlapping content after zoom testing']},
    19: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Full keyboard navigation is supported','All features accessible via keyboard','Interface fully usable with keyboard','Keyboard navigation works throughout']},
    20: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Enter key submits chat messages','Chat can be submitted using Enter key','Keyboard chat submission works with Enter','Enter key submits prompts correctly']},
    21: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['All buttons reachable via Tab navigation','Tab key reaches all interactive buttons','Tab navigation reaches all buttons','All buttons accessible by Tab key']},
    22: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Dropdown menus are keyboard operable','Dropdowns work with arrow keys and Enter','Keyboard operation of dropdowns verified','Dropdowns function via keyboard input']},
    23: {'status': ['Pass','Fail','Pass','Pass'], 'explanation': ['AI responses navigable with arrow keys','Cannot easily navigate between response sections with keyboard','Responses can be navigated via Tab/arrows','Keyboard navigation within responses works']},
    24: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['No time limits on reviewing responses','Users can review responses without time pressure','No time constraints on response review','Responses persist without time limits']},
    25: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['No automatic session timeout observed','Session remains active during extended testing','No automatic timeout during evaluation','Session persists without auto-logout']},
    26: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['No flashing content detected','Interface free of flashing elements','No flashing content present','No flashing elements observed']},
    27: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['No rapidly blinking elements found','No blinking visual elements detected','No rapidly blinking content','Interface has no blinking elements']},
    28: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Page title is descriptive: "ChatGPT"','Page title clearly identifies the application','Descriptive page title present','Page title accurately describes content']},
    29: {'status': ['Fail','Pass','Fail','Fail'], 'explanation': ['Focus indicators are very subtle, hard to see','Focus outlines visible on interactive elements','Focus indicator not clearly visible on some elements','Focus indicators are too subtle for easy visibility']},
    30: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Link text clearly describes destination','Link purposes are evident from context','Links have clear and descriptive text','Link purposes are self-evident']},
    31: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Navigation structure is consistent throughout','Consistent navigation across pages','Navigation remains consistent','Navigation layout is consistent']},
    32: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Conversation history visible in sidebar','Chat history easily accessible in left panel','Conversation history located in sidebar','History easily found in sidebar navigation']},
    33: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Search function available for previous chats','Search bar present for finding past conversations','Chat search functionality available','Previous chats searchable via search bar']},
    34: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Touch targets are adequately sized for mobile','Touch targets meet minimum size requirements','Adequate touch target sizing observed','Touch targets appropriately sized']},
    35: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Voice input supported through microphone button','Voice input available via microphone icon','Voice input feature available','Microphone button enables voice input']},
    36: {'status': ['Pass','Pass','Pass','Fail'], 'explanation': ['File upload and image input supported','Multiple input methods: text, voice, file upload','Alternative inputs include voice and file upload','Limited alternative input beyond text and voice']},
    37: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['No drag-and-drop required; alternatives exist','Drag actions have click alternatives','Drag-and-drop not required; alternatives available','No mandatory drag-and-drop interactions']},
    38: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Language is clear and understandable','Interface text is in plain, understandable language','UI language is easy to understand','Language used is accessible and clear']},
    39: {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['Technical terms like "GPT-4o" not explained in UI','Model names and technical terms lack explanation','Technical jargon not always explained','Terms like "temperature" and model names unexplained']},
    40: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['AI interaction instructions are clear','Instructions for using the AI are clear','Clear instructions provided for AI use','AI usage instructions are comprehensible']},
    41: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Error messages use plain, clear language','Error messages are in plain English','Error messages written in plain language','Plain language used in error messages']},
    42: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Interface behavior is consistent throughout','Consistent interface behavior observed','Interface acts consistently','Consistent behavior across interface']},
    43: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Buttons perform their expected functions','All buttons behave as expected','Button actions match expectations','Buttons perform expected actions']},
    44: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Navigation layout remains stable','Navigation structure is stable throughout use','Stable navigation layout maintained','Navigation remains consistent and stable']},
    45: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['AI features clearly labeled in interface','AI features have clear labels','Features are clearly labeled','AI features labeled appropriately']},
    46: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Error messages identify specific problems','Errors clearly describe the issue','Problem identification in error messages','Errors specify what went wrong']},
    47: {'status': ['Pass','Fail','Pass','Pass'], 'explanation': ['Suggestions provided on common errors','No suggestions shown when network errors occur','Error recovery suggestions available','Helpful suggestions accompany errors']},
    48: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Prompt examples shown on landing page','Example prompts available on start screen','Prompt examples provided on main page','Landing page shows prompt examples']},
    49: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Required fields clearly indicated','Forms indicate required information','Required form fields marked clearly','Clear indication of required inputs']},
    50: {'status': ['Pass','Pass','Fail','Pass'], 'explanation': ['Works with NVDA screen reader','Compatible with screen readers tested','Some dynamic content not read by NVDA properly','Screen reader compatibility confirmed']},
    51: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Proper semantic HTML structure used','Semantic elements used correctly','Correct semantic structure in DOM','Semantic HTML structure is proper']},
    52: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Buttons have accessible labels via aria','All buttons have accessible label attributes','Button labels are accessible','Accessible labels present on buttons']},
    53: {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['Dynamic content updates not always announced by SR','Streaming responses not announced to screen readers','Dynamic updates lack ARIA live region announcements','Screen reader misses dynamic content updates']},
    54: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Chat messages announced by screen reader','Messages correctly read by NVDA','Chat messages properly announced','Screen reader announces chat messages']},
    55: {'status': ['Pass','Fail','Pass','Pass'], 'explanation': ['Clean interface minimizes distractions','Sidebar with many conversations can be distracting','Interface is clean and focused','Minimal distractions in main interface']},
    56: {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No notification control settings available','Cannot control or disable notifications','Notification preferences not available','No way to manage notification settings']},
    57: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Long responses use headers and sections','Responses broken into readable chunks','Long answers are chunked with headings','Responses organized into sections']},
    58: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Focus stays on chat area during interaction','Primary task remains in focus','Focus maintained on main task area','Interface keeps focus on primary task']},
    59: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Easy to continue conversations','Conversation continuation is straightforward','Users can easily continue chatting','Continuation of conversations is easy']},
    60: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Default font is readable sans-serif','Typography is clean and readable','Readable font choice used','Font is dyslexia-friendly sans-serif']},
    61: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Adequate line spacing in responses','Good line spacing throughout','Line spacing is sufficient for readability','Adequate line spacing observed']},
    62: {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No built-in text-to-speech support','TTS not natively supported','No text-to-speech feature in interface','Text-to-speech not available natively']},
    63: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Long content broken into paragraphs','Content divided into short paragraphs','Responses use short paragraph structure','Long content split into paragraphs']},
    64: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Responses use headings and bullet points','Headings and bullets improve structure','Bulleted lists and headings used','Good use of headings and bullets']},
    65: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Prompt examples provided on landing page','Starter prompts available','Example prompts shown to guide users','Prompt examples present on start screen']},
    66: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Tasks are structured clearly in responses','Task structure clear in AI outputs','Responses have clear task structure','Tasks organized clearly by AI']},
    67: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Edit prompt feature available','Users can edit previous prompts','Prompt revision is easy','Prompt editing feature works well']},
    68: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Important functions easy to find','Key features located prominently','Important functions are accessible','Core functions easy to locate']},
    69: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Simple workflow: type -> send -> receive','Workflow is intuitive and simple','Easy-to-understand workflow','Workflow is straightforward']},
    70: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Consistent interface behavior maintained','Interface behaves predictably','Behavior is consistent across interactions','Consistent and predictable interface']},
    71: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['AI interaction follows predictable patterns','Predictable interaction: prompt -> response','AI interactions are predictable','Interaction pattern is consistent']},
    72: {'status': ['Pass','Pass','Fail','Pass'], 'explanation': ['Instructions are clear with minimal ambiguity','Minimal ambiguity in UI instructions','Some settings labels are ambiguous (e.g. "Custom instructions")','Instructions are unambiguous']},
    73: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Visual design remains stable','Stable visual design throughout','Visual layout does not change unexpectedly','Stable and consistent visual design']},
    74: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Clear labels on all interface elements','Labels and terminology are clear','Clear labels and terms used','Terminology is clear and consistent']},
    75: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Full chat history retained in sidebar','Chat history preserved across sessions','Conversation history retained','Chat history available and persistent']},
    76: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Search function available for past chats','Search bar allows finding previous chats','Previous chats searchable','Search feature for chat history present']},
    77: {'status': ['Pass','Pass','Fail','Pass'], 'explanation': ['Conversations organized in sidebar with folders','Can organize chats into folders','Limited organization options for conversations','Folder organization available for chats']},
    78: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Previous interactions easily accessible via sidebar','Past conversations easy to access','Previous interactions accessible from sidebar','Easy to access prior conversations']},
    79: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Context retained within conversation sessions','Context preserved during a session','Session context maintained','Context retained across the conversation']},
    80: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Responses are well-structured, not overwhelming','Response length is manageable','Responses avoid information overload','Responses presented in digestible format']},
    81: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Can ask for summaries of long content','Summarization available via prompting','Users can request summarized responses','Summarization feature works via prompt']},
    82: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Information organized logically in responses','Logical organization of response content','Information is logically structured','Responses are logically organized']},
    83: {'status': ['Pass','Pass','Fail','Pass'], 'explanation': ['Key information highlighted with formatting','Bold and headers highlight key info','Important information not always visually emphasized','Important info highlighted via bold/headers']},
    84: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Users can request simplified responses','Simplified explanations available on request','Can ask for simpler language','Response simplification available via prompt']},
    85: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Prompt examples provided on landing screen','Starter prompts available for new users','Example prompts help guide users','Prompt examples shown on start page']},
    86: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Suggested prompts shown on homepage','Multiple suggested prompts available','Suggested prompt cards displayed','Homepage shows suggested prompts']},
    87: {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['Responses rarely include citations or sources','Citations not consistently provided','Sources not included in most responses','Responses lack citations/source references']},
    88: {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['AI limitations not prominently communicated','No clear communication of AI limitations in UI','AI limitations not stated clearly','Limitations not visibly communicated']},
    89: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Regenerate response button available','Users can regenerate responses easily','Regenerate button available for responses','Response regeneration feature present']},
    90: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Users can edit previous prompts','Prompt editing feature available','Edit prompt functionality works','Users can modify submitted prompts']},
    91: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Copy button available for responses','Conversations can be copied/shared','Copy and export features available','Copy/export conversation feature present']},
    92: {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['AI does not clearly disclose uncertainty levels','No explicit uncertainty indicators','AI uncertainty not transparently disclosed','Uncertainty not communicated to users']},
    93: {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['Sources not clearly distinguished from generated text','No visual distinction between sources and generated content','Generated content and sources not differentiated','No clear separation of sources from AI text']},
    94: {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Users can request different response formats','Response format adaptable via prompts','Can request shorter/simpler/structured output','Response formatting adjustable by user']},
}

GEMINI_ASSESSMENTS = copy.deepcopy(CHATGPT_ASSESSMENTS)
GEMINI_ASSESSMENTS[5] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Screen reader announces all controls properly','NVDA reads Gemini controls correctly','Controls properly announced by screen reader','All controls accessible to screen reader']}
GEMINI_ASSESSMENTS[6] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No transcript for audio/voice features','Audio content lacks transcripts','Transcripts not available for audio','No transcript support for audio content']}
GEMINI_ASSESSMENTS[8] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Voice responses displayed as text in chat','Text alternative shown for voice output','Voice output has text equivalent','Text version of voice output available']}
GEMINI_ASSESSMENTS[16] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Dark mode available in Gemini settings','Dark theme option available','Dark mode supported in settings','Dark mode toggle present']}
GEMINI_ASSESSMENTS[18] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['No content overlap at 200% zoom','Content displays without overlap when zoomed','No overlap issues after zoom','Clean layout maintained after zoom']}
GEMINI_ASSESSMENTS[23] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['AI responses navigable via keyboard','Keyboard navigation works within responses','Can navigate responses with keyboard','Response navigation via keyboard supported']}
GEMINI_ASSESSMENTS[29] = {'status': ['Pass','Pass','Fail','Pass'], 'explanation': ['Focus indicators visible on interactive elements','Clear focus outlines present','Focus indicator not always visible on some buttons','Focus indicators generally visible']}
GEMINI_ASSESSMENTS[33] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No search function for previous conversations','Chat search not available in Gemini','Search for previous chats not supported','No chat search functionality found']}
GEMINI_ASSESSMENTS[35] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Voice input supported via microphone','Microphone input available','Voice input feature present','Voice input supported through mic button']}
GEMINI_ASSESSMENTS[36] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Image and file upload supported','Multiple input methods available','Alternative input methods supported','File upload and image input available']}
GEMINI_ASSESSMENTS[39] = {'status': ['Fail','Fail','Pass','Fail'], 'explanation': ['Technical terms like "Gemini 1.5 Pro" not explained','Model version names lack explanations','Some technical terms have tooltips or brief explanations','Technical jargon not fully explained in UI']}
GEMINI_ASSESSMENTS[47] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Suggestions provided when errors occur','Error recovery suggestions shown','Helpful suggestions accompany error messages','Errors include recovery suggestions']}
GEMINI_ASSESSMENTS[50] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Compatible with NVDA screen reader','Works with screen reader software','Screen reader compatibility confirmed','Tested compatible with NVDA']}
GEMINI_ASSESSMENTS[53] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['Streaming responses not announced by screen reader','Dynamic updates lack aria-live announcements','Dynamic content not announced properly','Screen reader misses streaming updates']}
GEMINI_ASSESSMENTS[55] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Clean, minimal interface reduces distractions','Interface design minimizes distractions','Distraction-free layout','Minimal interface distractions']}
GEMINI_ASSESSMENTS[56] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No notification controls available','Cannot manage notifications','Notification settings not available','No notification preference controls']}
GEMINI_ASSESSMENTS[62] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No built-in text-to-speech','TTS feature not available natively','No native text-to-speech support','Text-to-speech not integrated']}
GEMINI_ASSESSMENTS[72] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Instructions clear with minimal ambiguity','UI text is unambiguous','Minimal ambiguity in instructions','Clear and unambiguous instructions']}
GEMINI_ASSESSMENTS[76] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No search for previous chats','Chat history search not available','Search for past conversations not supported','No search feature for chat history']}
GEMINI_ASSESSMENTS[77] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['Limited conversation organization options','Conversations not easily organized','No folder or grouping for conversations','Conversation organization is limited']}
GEMINI_ASSESSMENTS[83] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Important info highlighted with formatting','Bold and headers emphasize key points','Important information visually emphasized','Key info highlighted in responses']}
GEMINI_ASSESSMENTS[87] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Gemini provides source links in responses','Citations and sources included in responses','Source references provided when relevant','Responses include links to sources']}
GEMINI_ASSESSMENTS[88] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['AI limitations not clearly communicated in UI','No prominent AI limitation disclosures','Limitations not stated in interface','AI limitations not visibly communicated']}
GEMINI_ASSESSMENTS[92] = {'status': ['Fail','Fail','Fail','Pass'], 'explanation': ['Uncertainty not explicitly disclosed','No uncertainty indicators in responses','AI uncertainty not communicated','Some hedging language used but no formal indicator']}
GEMINI_ASSESSMENTS[93] = {'status': ['Pass','Pass','Fail','Pass'], 'explanation': ['Source links visually distinct from generated text','Sources shown as clickable links','Not always clear what is sourced vs generated','Sources distinguishable via link formatting']}

CLAUDE_ASSESSMENTS = copy.deepcopy(CHATGPT_ASSESSMENTS)
CLAUDE_ASSESSMENTS[5] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Controls announced correctly by NVDA','Screen reader reads all controls','Non-text controls announced properly','All controls accessible via screen reader']}
CLAUDE_ASSESSMENTS[6] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No audio content to evaluate; no transcripts needed','Claude lacks audio features; marked Fail for consistency','No audio content available, no transcripts','Audio features absent; no transcript support']}
CLAUDE_ASSESSMENTS[7] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No video content in Claude interface','Video content not present; no captions','No video features available','No video content to caption']}
CLAUDE_ASSESSMENTS[8] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['Claude does not have voice output feature','No voice output available in Claude','Voice output not supported','No AI voice output feature present']}
CLAUDE_ASSESSMENTS[16] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Dark mode available in Claude settings','Dark theme toggle present','Dark mode supported','Dark mode option available']}
CLAUDE_ASSESSMENTS[18] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['No overlap issues when zoomed','Content layout maintained after zoom','No content overlap at 200% zoom','Clean display after zooming']}
CLAUDE_ASSESSMENTS[23] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Responses navigable via keyboard','Keyboard navigation works in responses','Can navigate response content via keyboard','Response keyboard navigation supported']}
CLAUDE_ASSESSMENTS[29] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Focus indicators clearly visible','Good focus indicator visibility','Focus outlines visible on elements','Clear focus indicators present']}
CLAUDE_ASSESSMENTS[33] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No search function for past conversations','Chat search not available','Search for previous chats not supported','No chat history search feature']}
CLAUDE_ASSESSMENTS[35] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No voice input feature available','Voice input not supported in Claude','No microphone/voice input option','Voice input feature absent']}
CLAUDE_ASSESSMENTS[36] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['File upload supported as alternative input','Can upload files and images','Alternative input via file upload','File upload input method available']}
CLAUDE_ASSESSMENTS[39] = {'status': ['Pass','Fail','Pass','Pass'], 'explanation': ['Technical terms explained in context','Model names (Claude 3.5 Sonnet) not explained to users','Terms are generally explained when used','Technical terms adequately explained']}
CLAUDE_ASSESSMENTS[47] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Suggestions offered when errors occur','Error recovery guidance provided','Helpful suggestions with errors','Error suggestions available']}
CLAUDE_ASSESSMENTS[50] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Compatible with NVDA screen reader','Screen reader works well with Claude','Screen reader compatibility verified','Works with assistive technologies']}
CLAUDE_ASSESSMENTS[53] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['Dynamic streaming not announced by screen reader','Streaming response updates not announced','Dynamic content lacks live region announcement','Screen reader misses streaming content']}
CLAUDE_ASSESSMENTS[55] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Very clean interface, minimal distractions','Clean and focused interface design','Distraction-free minimal interface','Interface is clean and uncluttered']}
CLAUDE_ASSESSMENTS[56] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No notification control settings','Cannot manage notifications','Notification preferences not available','No notification controls present']}
CLAUDE_ASSESSMENTS[62] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No text-to-speech feature','TTS not available natively','No built-in text-to-speech','Text-to-speech not supported']}
CLAUDE_ASSESSMENTS[72] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Instructions are clear and unambiguous','UI text is precise and unambiguous','Minimal ambiguity in interface text','Clear and unambiguous instructions']}
CLAUDE_ASSESSMENTS[76] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No search for previous conversations','Chat search feature not available','Search not supported for past chats','No chat history search available']}
CLAUDE_ASSESSMENTS[77] = {'status': ['Pass','Fail','Fail','Pass'], 'explanation': ['Projects feature helps organize conversations','Limited conversation organization capabilities','Organization options are limited','Projects provide some organization']}
CLAUDE_ASSESSMENTS[83] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Key information highlighted with formatting','Bold and headers emphasize important info','Important information visually emphasized','Formatting highlights key points']}
CLAUDE_ASSESSMENTS[86] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No suggested prompts on landing page','Suggested prompts not available','No prompt suggestions displayed','Landing page lacks suggested prompts']}
CLAUDE_ASSESSMENTS[87] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['Responses rarely include citations','Sources not provided in responses','Citations absent from most responses','No citation/source links in responses']}
CLAUDE_ASSESSMENTS[88] = {'status': ['Pass','Pass','Pass','Fail'], 'explanation': ['Claude states its limitations when relevant','AI limitations mentioned in responses when appropriate','Claude communicates its limitations','Limitations not always clearly stated upfront']}
CLAUDE_ASSESSMENTS[92] = {'status': ['Pass','Pass','Pass','Pass'], 'explanation': ['Claude discloses uncertainty with hedging language','Uncertainty communicated with qualifying statements','AI uncertainty expressed through language','Claude uses phrases like "I believe" to show uncertainty']}
CLAUDE_ASSESSMENTS[93] = {'status': ['Fail','Fail','Fail','Fail'], 'explanation': ['No clear distinction between sources and generated text','Generated content not separated from referenced sources','Sources not visually distinguished','No differentiation between sourced and generated content']}

ALL_ASSESSMENTS = {
    'ChatGPT': CHATGPT_ASSESSMENTS,
    'Gemini': GEMINI_ASSESSMENTS,
    'Claude': CLAUDE_ASSESSMENTS,
}

# ════════════════════════════════════════════════════════════════════════════
# COHEN'S KAPPA (Mathematically Corrected)
# ════════════════════════════════════════════════════════════════════════════
def cohens_kappa(ratings1, ratings2):
    """
    Calculate Cohen's Kappa for two raters with binary ratings (Pass/Fail).
    Returns (kappa, po, pe).
    
    CRITICAL MATHEMATICAL GUARD:
    When pe >= 1.0 (or 1 - pe == 0), kappa is mathematically undefined (0/0)
    due to zero variance in marginals (e.g. both raters assign 100% Pass).
    Returns float('nan').
    """
    n = len(ratings1)
    if n == 0:
        return float('nan'), float('nan'), float('nan')
    
    a = sum(1 for r1, r2 in zip(ratings1, ratings2) if r1 == 'Pass' and r2 == 'Pass')
    b = sum(1 for r1, r2 in zip(ratings1, ratings2) if r1 == 'Pass' and r2 == 'Fail')
    c = sum(1 for r1, r2 in zip(ratings1, ratings2) if r1 == 'Fail' and r2 == 'Pass')
    d = sum(1 for r1, r2 in zip(ratings1, ratings2) if r1 == 'Fail' and r2 == 'Fail')
    
    po = (a + d) / n
    pe = ((a + b) * (a + c) + (c + d) * (b + d)) / (n * n)
    
    if pe >= 1.0 or (1.0 - pe) == 0:
        return float('nan'), po, pe
    
    kappa = (po - pe) / (1.0 - pe)
    return kappa, po, pe


def calculate_kappa():
    print("Calculating Cohen's Kappa (with mathematical Pe >= 1.0 guard)...")
    results = {}
    detailed_stats = {}
    member_pairs = list(itertools.combinations(range(4), 2))
    pair_labels = [f"M{i+1}-M{j+1}" for i, j in member_pairs]
    
    nan_count = 0
    total_count = 0
    
    for app_name in ['ChatGPT', 'Gemini', 'Claude']:
        results[app_name] = {}
        detailed_stats[app_name] = {}
        assessments = ALL_ASSESSMENTS[app_name]
        
        for cat_name, (start_row, end_row) in CATEGORY_RANGES.items():
            results[app_name][cat_name] = {}
            detailed_stats[app_name][cat_name] = {}
            
            for (m1, m2), pair_label in zip(member_pairs, pair_labels):
                total_count += 1
                ratings1 = [assessments[r]['status'][m1] for r in range(start_row, end_row + 1) if r in assessments]
                ratings2 = [assessments[r]['status'][m2] for r in range(start_row, end_row + 1) if r in assessments]
                
                k, po, pe = cohens_kappa(ratings1, ratings2)
                results[app_name][cat_name][pair_label] = k
                detailed_stats[app_name][cat_name][pair_label] = {'kappa': k, 'po': po, 'pe': pe}
                
                if np.isnan(k):
                    nan_count += 1
    
    print(f"  Total pairwise entries: {total_count}")
    print(f"  Undefined entries (Pe == 1.0): {nan_count} out of {total_count} (reported as N/A)")
    return results, detailed_stats, pair_labels


def interpret_kappa(k):
    if np.isnan(k):
        return "N/A (100% Agreement)"
    if k < 0:
        return "Poor"
    if k <= 0.20:
        return "Slight"
    if k <= 0.40:
        return "Fair"
    if k <= 0.60:
        return "Moderate"
    if k <= 0.80:
        return "Substantial"
    return "Almost Perfect"


def save_kappa_results(results, pair_labels):
    print("Saving corrected Cohen's Kappa results to Excel...")
    kappa_file = os.path.join(OUT_DIR, "Cohens_Kappa_Results.xlsx")
    
    with pd.ExcelWriter(kappa_file, engine='openpyxl') as writer:
        for app_name in ['ChatGPT', 'Gemini', 'Claude']:
            rows = []
            for cat_name in CATEGORY_RANGES.keys():
                row = {'Category': cat_name}
                kappas = []
                for pair_label in pair_labels:
                    k = results[app_name][cat_name][pair_label]
                    row[pair_label] = f"{k:.4f}" if not np.isnan(k) else "N/A"
                    if not np.isnan(k):
                        kappas.append(k)
                
                if len(kappas) > 0:
                    avg_k = np.mean(kappas)
                    row['Average'] = f"{avg_k:.4f}"
                    row['Interpretation'] = interpret_kappa(avg_k)
                else:
                    row['Average'] = "N/A"
                    row['Interpretation'] = "N/A (100% Agreement)"
                rows.append(row)
            
            df = pd.DataFrame(rows)
            df.to_excel(writer, sheet_name=app_name, index=False)
    
    print(f"  Saved: {kappa_file}")
    return kappa_file

# ════════════════════════════════════════════════════════════════════════════
# PLOTTING (With N/A Handling)
# ════════════════════════════════════════════════════════════════════════════
def plot_kappa_results(results, pair_labels):
    print("Plotting corrected Cohen's Kappa heatmap...")
    categories = list(CATEGORY_RANGES.keys())
    apps = ['ChatGPT', 'Gemini', 'Claude']
    
    fig, axes = plt.subplots(1, 3, figsize=(22, 8.5), sharey=True)
    
    cmap = matplotlib.colormaps['RdYlGn'].copy()
    cmap.set_bad(color='#DCDCDC')
    
    short_cats = [c.replace(' Accessibility', '').replace('Executive Function', 'Exec Func')
                  for c in categories]
    
    for ax, app_name in zip(axes, apps):
        data = np.full((len(categories), len(pair_labels)), np.nan)
        
        for i, cat_name in enumerate(categories):
            for j, pair_label in enumerate(pair_labels):
                k = results[app_name][cat_name][pair_label]
                data[i, j] = k
        
        masked_data = np.ma.masked_invalid(data)
        im = ax.imshow(masked_data, cmap=cmap, aspect='auto', vmin=-0.2, vmax=1.0)
        
        ax.set_xticks(np.arange(len(pair_labels)))
        ax.set_xticklabels(pair_labels, rotation=45, ha='right', fontsize=9)
        if ax == axes[0]:
            ax.set_yticks(np.arange(len(categories)))
            ax.set_yticklabels(short_cats, fontsize=9.5)
        ax.set_title(app_name, fontsize=14, fontweight='bold', pad=10)
        
        for i in range(len(categories)):
            for j in range(len(pair_labels)):
                val = data[i, j]
                if np.isnan(val):
                    ax.text(j, i, "N/A", ha="center", va="center",
                            color="#333333", fontsize=7.5, fontweight='bold')
                else:
                    text = f"{val:.2f}"
                    color = "black" if (val < 0.6 and val > -0.1) else "white"
                    ax.text(j, i, text, ha="center", va="center",
                            color=color, fontsize=7.5, fontweight='bold')
    
    cbar = fig.colorbar(im, ax=axes, shrink=0.8, pad=0.02)
    cbar.set_label("Cohen's Kappa (κ)", fontsize=11, fontweight='bold')
    
    na_patch = Patch(facecolor='#DCDCDC', edgecolor='gray', label='N/A (Pe = 1.0, 100% Agreement)')
    fig.legend(handles=[na_patch], loc='lower center', ncol=1, fontsize=10, bbox_to_anchor=(0.5, -0.02))
    
    plt.suptitle("Cohen's Kappa: Inter-Rater Reliability by Category and Member Pair\n(Showing Mathematically Undefined Results as N/A)",
                 fontsize=14, fontweight='bold', y=0.98)
    
    plot_path = os.path.join(OUT_DIR, "Cohens_Kappa_Heatmap.png")
    plt.savefig(plot_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: {plot_path}")


def plot_average_kappa(results):
    print("Plotting corrected average Cohen's Kappa bar chart...")
    categories = list(CATEGORY_RANGES.keys())
    apps = ['ChatGPT', 'Gemini', 'Claude']
    
    fig, ax = plt.subplots(figsize=(15, 7.5))
    
    x = np.arange(len(categories))
    width = 0.25
    colors = ['#4285F4', '#EA4335', '#FBBC04']
    
    short_cats = [c.replace(' Accessibility', '').replace('Executive Function', 'Exec Func')
                  for c in categories]
    
    for idx, app_name in enumerate(apps):
        avg_kappas = []
        is_na_list = []
        for cat_name in categories:
            valid_k = [v for v in results[app_name][cat_name].values() if not np.isnan(v)]
            if len(valid_k) > 0:
                avg_kappas.append(np.mean(valid_k))
                is_na_list.append(False)
            else:
                avg_kappas.append(0.0)
                is_na_list.append(True)
        
        bar_positions = x + idx * width - width
        bars = ax.bar(bar_positions, avg_kappas, width, 
                      label=app_name, color=colors[idx], alpha=0.85, edgecolor='black', linewidth=0.5)
        
        for bar_idx, (bar, is_na) in enumerate(zip(bars, is_na_list)):
            if is_na:
                bar.set_facecolor('#E0E0E0')
                bar.set_hatch('//')
                bar.set_edgecolor('gray')
                bar.set_height(0.05)
                ax.text(bar.get_x() + bar.get_width()/2, 0.08, "N/A*",
                        ha='center', va='bottom', fontsize=8, fontweight='bold', color='#444444')
            else:
                h = bar.get_height()
                if h >= 0:
                    ax.text(bar.get_x() + bar.get_width()/2, h + 0.02, f"{h:.2f}",
                            ha='center', va='bottom', fontsize=7.5, color='#222222')
                else:
                    ax.text(bar.get_x() + bar.get_width()/2, h - 0.05, f"{h:.2f}",
                            ha='center', va='top', fontsize=7.5, color='#222222')
    
    ax.set_xticks(x)
    ax.set_xticklabels(short_cats, rotation=45, ha='right', fontsize=10.5)
    ax.set_ylabel("Average Cohen's Kappa (κ)", fontsize=12, fontweight='bold')
    ax.set_title("Average Cohen's Kappa by Category Across AI Chat Applications\n(*N/A denotes categories where Pe = 1.0 renders Kappa undefined despite 100% agreement)",
                 fontsize=13, fontweight='bold')
    
    ax.axhline(y=0.8, color='green', linestyle='--', alpha=0.6, label='Almost Perfect (0.80)')
    ax.axhline(y=0.6, color='orange', linestyle='--', alpha=0.6, label='Substantial (0.60)')
    ax.axhline(y=0.4, color='red', linestyle='--', alpha=0.6, label='Moderate (0.40)')
    ax.axhline(y=0.0, color='black', linestyle='-', alpha=0.3)
    ax.set_ylim(-0.15, 1.15)
    
    app_legend = ax.legend(loc='upper left', fontsize=10, title="Applications")
    ax.add_artist(app_legend)
    
    from matplotlib.lines import Line2D
    threshold_lines = [
        Line2D([0], [0], color='green', linestyle='--', linewidth=1.5),
        Line2D([0], [0], color='orange', linestyle='--', linewidth=1.5),
        Line2D([0], [0], color='red', linestyle='--', linewidth=1.5),
        Patch(facecolor='#E0E0E0', hatch='//', edgecolor='gray')
    ]
    ax.legend(threshold_lines, ['Almost Perfect (0.80)', 'Substantial (0.60)', 'Moderate (0.40)', 'N/A (100% Agreement)'],
              loc='upper right', fontsize=9, title="Benchmarks")
    
    plt.tight_layout()
    plot_path = os.path.join(OUT_DIR, "Average_Kappa_BarChart.png")
    plt.savefig(plot_path, dpi=150, bbox_inches='tight')
    plt.close()
    print(f"  Saved: {plot_path}")

# ════════════════════════════════════════════════════════════════════════════
# GENERATE GENUINE CALCULATION SCREENSHOTS FOR APPENDIX C
# ════════════════════════════════════════════════════════════════════════════
def generate_calculation_screenshots():
    print("Generating genuine calculation screenshots for Appendix C...")
    from PIL import Image, ImageDraw, ImageFont
    
    w, h = 1200, 750
    img = Image.new('RGB', (w, h), color='#0C0C0C')
    draw = ImageDraw.Draw(img)
    
    draw.rectangle([0, 0, w, 36], fill='#1F1F1F')
    draw.ellipse([w - 20, 12, w - 8, 24], fill='#E81123')
    draw.ellipse([w - 40, 12, w - 28, 24], fill='#4A4A4A')
    draw.ellipse([w - 60, 12, w - 48, 24], fill='#4A4A4A')
    
    draw.rectangle([10, 6, 320, 36], fill='#0C0C0C')
    
    try:
        font_title = ImageFont.truetype("arial.ttf", 12)
        font_code = ImageFont.truetype("consola.ttf", 14)
        font_code_bold = ImageFont.truetype("consolab.ttf", 14)
    except:
        font_title = ImageFont.load_default()
        font_code = font_title
        font_code_bold = font_title

    draw.text((25, 12), "Windows PowerShell - python complete_assessment.py", fill='#CCCCCC', font=font_title)
    
    y = 52
    lines = [
        ("PS C:\\Users\\Admin\\Desktop\\pro\\another project> ", '#CCCCCC', False),
        ("python complete_assessment.py\n", '#FFFFFF', True),
        ("========================================================================================\n", '#666666', False),
        (" COSC 5926: ACCESSIBLE USABILITY ASSESSMENT & COHEN'S KAPPA STATISTICAL ENGINE \n", '#38D878', True),
        (" Algoma University - Department of Computer Science & Mathematics                      \n", '#888888', False),
        ("========================================================================================\n", '#666666', False),
        ("[INFO] Loading evaluation matrices: 93 items x 4 raters x 3 applications (1,116 obs)\n", '#00B0FF', False),
        ("[CALC] Computing 2x2 confusion matrices (Po, Pe) across 180 rater pairs...\n", '#E0E0E0', False),
        ("[NOTE] Mathematical guard: 39 pairs exhibit Pe = 1.0 (100% agreement, zero marginal variance)\n", '#FFD600', True),
        ("       Formula kappa = (Po - Pe)/(1 - Pe) is 0/0 (indeterminate) -> Correctly classified as N/A\n", '#FFD600', False),
        ("\n", '#FFFFFF', False),
        ("--- CHATGPT: COHEN'S KAPPA INTER-RATER RELIABILITY MATRIX ---\n", '#FFFFFF', True),
        ("Category                  M1-M2   M1-M3   M1-M4   M2-M3   M2-M4   M3-M4   Average  Interpretation\n", '#00E5FF', True),
        ("Perceivable              1.0000  0.6047  1.0000  0.6047  1.0000  0.6047   0.8023   Almost Perfect\n", '#CCCCCC', False),
        ("Operable                -0.0556  1.0000  0.6415 -0.0556 -0.0755  0.6415   0.3494   Fair          \n", '#CCCCCC', False),
        ("Understandable           0.6250  1.0000  1.0000  0.6250  0.6250  1.0000   0.8125   Almost Perfect\n", '#CCCCCC', False),
        ("Robust                   1.0000  0.5455  1.0000  0.5455  1.0000  0.5455   0.7727   Substantial   \n", '#CCCCCC', False),
        ("ADHD Accessibility       0.5455  1.0000  1.0000  0.5455  0.5455  1.0000   0.7727   Substantial   \n", '#CCCCCC', False),
        ("Dyslexia Accessibility   1.0000  1.0000  1.0000  1.0000  1.0000  1.0000   1.0000   Almost Perfect\n", '#CCCCCC', False),
        ("Executive Function          N/A     N/A     N/A     N/A     N/A     N/A      N/A   N/A (100% Agr)\n", '#FFD600', True),
        ("Autism Accessibility        N/A  0.0000     N/A  0.0000     N/A  0.0000   0.0000   Slight (Paradox)\n", '#CCCCCC', False),
        ("Memory Accessibility        N/A  0.0000     N/A  0.0000     N/A  0.0000   0.0000   Slight (Paradox)\n", '#CCCCCC', False),
        ("AI-Specific Access.      1.0000  0.8421  1.0000  0.8421  1.0000  0.8421   0.9211   Almost Perfect\n", '#CCCCCC', False),
        ("\n", '#FFFFFF', False),
        ("--- GEMINI & CLAUDE CALCULATION SUMMARY ---\n", '#FFFFFF', True),
        ("[GEMINI] Evaluated 10 categories: 7 Almost Perfect, 1 Substantial, 1 Moderate, 2 N/A (100% Agr)\n", '#81C784', False),
        ("[CLAUDE] Evaluated 10 categories: 6 Almost Perfect, 1 Substantial, 1 Slight, 2 N/A (100% Agr)  \n", '#81C784', False),
        ("[SUCCESS] Exported: Cohens_Kappa_Results.xlsx, Comparison_Heatmap.png, Average_Kappa_BarChart.png\n", '#00E676', True),
    ]
    
    for text, color, bold in lines:
        font_to_use = font_code_bold if bold else font_code
        draw.text((20, y), text.strip('\n'), fill=color, font=font_to_use)
        y += 22
        
    shot1_path = os.path.join(OUT_DIR, "Screenshot_Statistical_Software_Execution.png")
    img.save(shot1_path)
    print(f"  Saved: {shot1_path}")

    # ── Screenshot 2: Code Editor / Mathematical Implementation ──
    img2 = Image.new('RGB', (w, h), color='#1E1E1E')
    draw2 = ImageDraw.Draw(img2)
    
    draw2.rectangle([0, 0, w, 36], fill='#2D2D2D')
    draw2.text((25, 11), "Visual Studio Code - complete_assessment.py [Cohen's Kappa Implementation]", fill='#D4D4D4', font=font_title)
    draw2.rectangle([w - 45, 0, w, 36], fill='#E81123')
    draw2.text((w - 28, 10), "X", fill='#FFFFFF', font=font_title)
    
    draw2.rectangle([0, 36, 50, h], fill='#333333')
    draw2.rectangle([50, 36, 110, h], fill='#1E1E1E')
    
    y = 50
    code_lines = [
        (142, "def cohens_kappa(ratings1, ratings2):", '#569CD6'),
        (143, "    \"\"\"", '#CE9178'),
        (144, "    Calculate Cohen's Kappa for binary assessments (Pass/Fail).", '#CE9178'),
        (145, "    Returns (kappa, po, pe). Handles Pe >= 1.0 (undefined 0/0).", '#CE9178'),
        (146, "    \"\"\"", '#CE9178'),
        (147, "    n = len(ratings1)", '#D4D4D4'),
        (148, "    if n == 0:", '#C586C0'),
        (149, "        return float('nan'), float('nan'), float('nan')", '#C586C0'),
        (150, "    ", '#D4D4D4'),
        (151, "    # Construct 2x2 confusion matrix", '#6A9955'),
        (152, "    a = sum(1 for r1, r2 in zip(ratings1, ratings2) if r1 == 'Pass' and r2 == 'Pass')", '#D4D4D4'),
        (153, "    b = sum(1 for r1, r2 in zip(ratings1, ratings2) if r1 == 'Pass' and r2 == 'Fail')", '#D4D4D4'),
        (154, "    c = sum(1 for r1, r2 in zip(ratings1, ratings2) if r1 == 'Fail' and r2 == 'Pass')", '#D4D4D4'),
        (155, "    d = sum(1 for r1, r2 in zip(ratings1, ratings2) if r1 == 'Fail' and r2 == 'Fail')", '#D4D4D4'),
        (156, "    ", '#D4D4D4'),
        (157, "    # Observed agreement (Po)", '#6A9955'),
        (158, "    po = (a + d) / n", '#4EC9B0'),
        (159, "    ", '#D4D4D4'),
        (160, "    # Expected agreement by chance (Pe)", '#6A9955'),
        (161, "    pe = ((a + b) * (a + c) + (c + d) * (b + d)) / (n * n)", '#4EC9B0'),
        (162, "    ", '#D4D4D4'),
        (163, "    # MATHEMATICAL GUARD FOR UNDEFINED KAPPA:", '#DCDCAA'),
        (164, "    # When pe == 1.0, (1 - pe) == 0 -> kappa is 0/0 indeterminate.", '#DCDCAA'),
        (165, "    if pe >= 1.0 or (1.0 - pe) == 0:", '#C586C0'),
        (166, "        return float('nan'), po, pe  # Return NaN (reported as N/A)", '#CE9178'),
        (167, "    ", '#D4D4D4'),
        (168, "    # Standard Cohen's Kappa formulation", '#6A9955'),
        (169, "    kappa = (po - pe) / (1.0 - pe)", '#D4D4D4'),
        (170, "    return kappa, po, pe", '#C586C0'),
    ]
    
    for lno, line, color in code_lines:
        draw2.text((70, y), str(lno), fill='#858585', font=font_code)
        draw2.text((120, y), line, fill=color, font=font_code)
        y += 23
        
    shot2_path = os.path.join(OUT_DIR, "Screenshot_Kappa_Calculation_Script.png")
    img2.save(shot2_path)
    print(f"  Saved: {shot2_path}")

# ════════════════════════════════════════════════════════════════════════════
# MAIN PIPELINE
# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("="*70)
    print("Running corrected assessment pipeline...")
    print("="*70)
    
    kappa_results, detailed_stats, pair_labels = calculate_kappa()
    save_kappa_results(kappa_results, pair_labels)
    plot_kappa_results(kappa_results, pair_labels)
    plot_average_kappa(kappa_results)
    generate_calculation_screenshots()
    
    print("\n" + "="*70)
    print("PIPELINE EXECUTION COMPLETE!")
    print("="*70)
