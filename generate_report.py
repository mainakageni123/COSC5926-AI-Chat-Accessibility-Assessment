"""
Generate the finalized Word document following the user's exact wording and report structure,
updating only the required sections: Table 4, Section 4.3 Kappa text, Section 5.4 discussion,
Section 6 conclusion, References [16]-[17], and Appendix C screenshots.
"""

import os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

OUT_DIR = r"c:\Users\Admin\Desktop\pro\another project"
REPORT_PATH = os.path.join(OUT_DIR, "COSC5926_Assignment1_Report.docx")

doc = Document()

style = doc.styles['Normal']
font = style.font
font.name = 'Times New Roman'
font.size = Pt(12)
style.paragraph_format.space_after = Pt(6)
style.paragraph_format.line_spacing = 1.5

for section in doc.sections:
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)

# ── Title Page ──
for _ in range(5):
    doc.add_paragraph('')

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Assignment 1')
run.font.size = Pt(26)
run.bold = True

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Assessing the Accessible Usability of AI Chat Applications')
run.font.size = Pt(18)
run.bold = True

doc.add_paragraph('')

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('COSC 5926: Advanced Human Computer Interaction')
run.font.size = Pt(14)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Department of Computer Science & Mathematics\nAlgoma University')
run.font.size = Pt(13)

for _ in range(4):
    doc.add_paragraph('')

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Submitted to: Professor Anjum')
run.font.size = Pt(12)

doc.add_paragraph('')

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('November 2026')
run.font.size = Pt(12)

doc.add_page_break()

# ── Table of Contents ──
doc.add_heading('Table of Contents', level=1)
toc = [
    '1. Introduction',
    '2. Methodology',
    '   2.1 Apps we tested',
    '   2.2 The checklist',
    '   2.3 Tools we used',
    '   2.4 How we did the evaluation',
    '3. What We Checked and What We Saw',
    '   3.1 The four WCAG principles',
    '   3.2 Cognitive accessibility',
    '   3.3 Accessibility problems specific to AI',
    '4. Results',
    '   4.1 Overall pass and fail counts',
    '   4.2 Results by category',
    '   4.3 How well we agreed with each other (Cohen\'s Kappa)',
    '5. Discussion',
    '   5.1 What the apps do well',
    '   5.2 Problems all three share',
    '   5.3 How the apps differ',
    '   5.4 What the agreement scores tell us',
    '   5.5 Limits of this study',
    '6. Conclusion',
    '7. References',
    '8. Appendices',
]
for item in toc:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(item)
    run.font.size = Pt(11)
    if not item.startswith('   '):
        run.bold = True

doc.add_page_break()

# ── 1. Introduction ──
doc.add_heading('1. Introduction', level=1)
doc.add_paragraph(
    'A huge number of people now lean on AI chat apps like ChatGPT, Google Gemini and Claude for writing, studying and work. '
    'That is exactly why accessibility matters here. Once an app becomes part of how people do their jobs and finish their coursework, '
    'a feature that only works if you can see the screen, use a mouse and read quickly is not a small flaw. It decides who gets to use the '
    'tool at all. We wanted to find out whether these apps pass that test.'
)
doc.add_paragraph(
    'Usability testing usually asks whether people can finish a task, how quickly they can do it, and whether they enjoy the experience [14]. '
    'Those questions still matter, but AI chat apps add a few new ones. A long, dense answer can overwhelm a reader. It is not always clear which '
    'parts were written by the AI or where the facts came from. And for people who are blind, have low vision, find it hard to use a mouse, or '
    'live with conditions like ADHD or dyslexia, the problem can be much simpler: a feature might just not work for them.'
)
doc.add_paragraph(
    'We chose to compare three apps instead of assessing just one because accessibility depends on how each platform is built. The apps offer '
    'similar services, but they do not necessarily give the same experience to someone using assistive technology. Putting them side by side '
    'let us see where one app does better or worse than the others, and whether some problems turn up in all three.'
)
doc.add_paragraph(
    'In this project we tested three of the best-known apps: ChatGPT, Gemini and Claude. We built a checklist of 93 items from three sources. '
    'The first is the Web Content Accessibility Guidelines (WCAG) 2.1 [2], which is the standard rulebook for making websites accessible. '
    'The second is our own checklist for cognitive accessibility, meaning things that affect attention, reading, planning and memory. '
    'The third is a short list of checks for problems that only come up with AI. Four of us went through every item for every app, each on our own, '
    'so that nobody\'s answers were influenced by anyone else\'s. That adds up to 93 items × 3 apps × 4 evaluators = 1,116 separate judgments. '
    'We then used a score called Cohen\'s Kappa [4] to see how much the four of us agreed.'
)
doc.add_paragraph(
    'What we wanted was simple: to see where each app does well, where it lets people down, and what could be changed so that more people can use '
    'these tools comfortably.'
)

# ── 2. Methodology ──
doc.add_heading('2. Methodology', level=1)

doc.add_heading('2.1 Apps we tested', level=2)
doc.add_paragraph('We picked three apps because they are well known and used by a lot of people:')
doc.add_paragraph('ChatGPT (OpenAI) [11]. Probably the most widely used of the three. It supports typing and voice (talking to it and hearing it talk back), file uploads and image creation.', style='List Bullet')
doc.add_paragraph('Google Gemini [12]. Google\'s assistant, built to work with Google\'s other services. You can give it text, images or voice, and its answers can include links to the sources it used.', style='List Bullet')
doc.add_paragraph('Claude (Anthropic) [13]. An assistant that aims to be helpful, safe and honest. It offers text chat and file upload, and it has a plain, uncluttered look.', style='List Bullet')

doc.add_heading('2.2 The checklist', level=2)
doc.add_paragraph(
    'The checklist has 93 items sorted into 10 categories, and those categories fall into three groups. '
    'The first group follows the WCAG 2.1 rulebook [2], which builds on WCAG 2.0 [1], and has 53 items:'
)

t_wcag = doc.add_table(rows=5, cols=3)
t_wcag.style = 'Light Shading Accent 1'
t_wcag.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, h in enumerate(['Category', 'What it covers', 'Items']):
    t_wcag.rows[0].cells[i].text = h
    t_wcag.rows[0].cells[i].paragraphs[0].runs[0].bold = True
wcag_rows = [
    ('Perceivable', 'Text descriptions for images, captions for audio and video, flexible layout, content that is easy to see', '17'),
    ('Operable', 'Using the keyboard, having enough time, avoiding flashing that can cause seizures, finding your way around, different ways of giving input', '19'),
    ('Understandable', 'Readable text, predictable behaviour, help with filling things in and fixing mistakes', '12'),
    ('Robust', 'Working properly with screen readers and other assistive tools', '5'),
]
for r, (c, w, itm) in enumerate(wcag_rows, 1):
    t_wcag.rows[r].cells[0].text = c
    t_wcag.rows[r].cells[1].text = w
    t_wcag.rows[r].cells[2].text = itm

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Group 1: WCAG 2.1 categories (53 items)')
run.italic = True
run.font.size = Pt(10)

doc.add_paragraph(
    'The second group is our cognitive accessibility checklist, with 25 items. It has five categories of five items each:'
)

t_cog = doc.add_table(rows=6, cols=2)
t_cog.style = 'Light Shading Accent 1'
t_cog.alignment = WD_TABLE_ALIGNMENT.CENTER
t_cog.rows[0].cells[0].text = 'Category'
t_cog.rows[0].cells[0].paragraphs[0].runs[0].bold = True
t_cog.rows[0].cells[1].text = 'Items'
t_cog.rows[0].cells[1].paragraphs[0].runs[0].bold = True
cog_rows = [
    ('ADHD accessibility', '5'),
    ('Dyslexia accessibility', '5'),
    ('Executive function accessibility (planning and organising tasks)', '5'),
    ('Autism-friendly accessibility', '5'),
    ('Memory accessibility', '5'),
]
for r, (c, itm) in enumerate(cog_rows, 1):
    t_cog.rows[r].cells[0].text = c
    t_cog.rows[r].cells[1].text = itm

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Group 2: Cognitive accessibility categories (25 items)')
run.italic = True
run.font.size = Pt(10)

doc.add_paragraph(
    'We developed the cognitive accessibility items for this evaluation as practical usability checks, informed by W3C guidance '
    'on cognitive and learning disabilities [5]. They are not clinically validated measures and were never meant to diagnose or measure '
    'any disability. The results should be read as observations about interface features, not as assessments of individual users.'
)
doc.add_paragraph(
    'The third group covers things specific to AI chat apps, such as how much mental effort an answer takes to read, how open the app '
    'is about what it can and cannot do, and how much control the user has over the answers. This group has 15 items and counts as one '
    'category in our results, called AI-Specific. Together the three groups give 53 + 25 + 15 = 93 items.'
)
doc.add_paragraph(
    'We used these three groups because checking WCAG alone would not give a complete picture of the user experience. An app can meet the '
    'basic technical requirements and still be hard to use for someone who struggles with concentration, reading or remembering information. '
    'AI chat apps also have features that traditional website checks do not fully cover, especially the way answers are generated and '
    'presented. Combining the three groups gave us a broader basis for comparing the apps.'
)

doc.add_heading('2.3 Tools we used', level=2)
doc.add_paragraph('Some things can be checked by software and some only by a person looking and trying things out [6][15], so we used a mix:')
t_list = [
    ('Manual inspection', 'Simply looking over each app\'s layout, design and features.'),
    ('WAVE [7]', 'A browser add-on that points out accessibility errors and warnings on a web page.'),
    ('axe DevTools [8]', 'Another browser tool. It checks a page against WCAG rules and explains what is wrong and how to fix it.'),
    ('NVDA [9]', 'A free screen reader for Windows. A screen reader reads the screen out loud for people who cannot see it, so this showed us what a blind user would hear.'),
    ('Keyboard-only testing', 'Doing everything with the keyboard (Tab, Enter, arrow keys, Escape) and no mouse.'),
    ('Colour Contrast Analyser [10]', 'Measures how strongly text stands out from its background. WCAG asks for at least 4.5 to 1 for normal text and 3 to 1 for large text.'),
    ('Browser zoom at 200%', 'Enlarging the page to double size to check that text can grow and the layout does not fall apart.'),
]
for name, desc in t_list:
    p = doc.add_paragraph(style='List Bullet')
    run = p.add_run(f'{name}. ')
    run.bold = True
    p.add_run(desc)

doc.add_heading('2.4 How we did the evaluation', level=2)
doc.add_paragraph(
    'Four group members (Member 1 to Member 4) each went through all 93 items for all three apps, working alone. For every item we recorded '
    'a Pass or a Fail and wrote a short reason for the decision. We did this in November 2026, using the latest public version of each app '
    'in Google Chrome on Windows computers.'
)
doc.add_paragraph(
    'Working alone mattered. Cohen\'s Kappa only means something if each score is one person\'s own opinion. If we had talked each item '
    'through and agreed on an answer as a group, there would be nothing left to measure.'
)
doc.add_paragraph(
    'Not every item was equally easy to judge. Checking whether a button can be reached with the keyboard is fairly straightforward. '
    'Deciding whether an explanation is clear enough, or whether an interface is easy to understand, takes more interpretation. We kept '
    'our individual results separate so that these differences would show up in the Kappa scores instead of being hidden inside one group '
    'decision. It also meant we sometimes saw the same interface differently, as the results in section 4.3 show.'
)

doc.add_page_break()

# ── 3. What We Checked and What We Saw ──
doc.add_heading('3. What We Checked and What We Saw', level=1)

doc.add_heading('3.1 The four WCAG principles', level=2)
doc.add_paragraph(
    'WCAG [2] is built on four ideas. We go through them one at a time, saying what each means in everyday terms and then what we found.'
)

doc.add_heading('Perceivable: can people take in the information?', level=3)
doc.add_paragraph(
    'This is about whether people can see, hear or otherwise notice what is on the screen. Pictures need written descriptions for people '
    'who cannot see them. Audio and video need text versions for people who cannot hear. Text has to stand out clearly from the background.'
)
doc.add_paragraph(
    'We checked descriptions and labels with WAVE, axe DevTools and NVDA, contrast with the Colour Contrast Analyser, and layout by zooming '
    'to 200%. Most of it held up. Most of the interface elements we examined had accessible labels (including alt text, the written description '
    'of an image), and the automated tools did not flag widespread missing descriptions. We used NVDA to hear how selected elements were '
    'announced, because a label being present does not mean it is useful. For contrast, we measured selected interface elements in the default '
    'and dark themes, not every possible screen, and the ones we checked met the WCAG ratios. A few of us noticed small layout glitches in '
    'ChatGPT\'s sidebar when zoomed in. Audio and video were the weak spot. In the voice and audio features we looked at, none of the three '
    'offered transcripts for audio or captions for video. ChatGPT and Gemini show a written version of their spoken replies, and Claude has '
    'no voice features, so there was nothing to compare.'
)
doc.add_paragraph(
    'What these checks showed is that passing a visual inspection does not automatically mean the same content is accessible to everyone. '
    'A page can look well organised to someone using a mouse and still be confusing for someone relying on a screen reader. That is why we '
    'looked at both how each app appears on screen and how its elements are presented through assistive technology.'
)

doc.add_heading('Operable: can people actually use it?', level=3)
doc.add_paragraph(
    'Not everyone uses a mouse. Some people rely on the keyboard alone, and some use other devices. This principle also covers having '
    'enough time, avoiding flashing that can set off seizures, and being able to find your way around.'
)
doc.add_paragraph(
    'The keyboard tests were the easy part. All the functions in our keyboard-only test could be completed without a mouse. We used Tab, '
    'Enter, arrow keys and Escape to reach and use the controls on our checklist: Enter sent a message, Tab reached the buttons we tried, '
    'and the drop-down menus opened from the keyboard. We saw no flashing or blinking content anywhere. The differences were in features. '
    'ChatGPT and Gemini let you speak your question, Claude does not, and only ChatGPT lets you search your old conversations. The focus '
    'outline is where we did not fully agree. This is the outline an app draws around the button you are on when you press Tab, so you '
    'can see where you are. Some of us found ChatGPT\'s outline too faint to follow comfortably. One of us kept losing track of where the '
    'focus was and had to move the mouse just to check. It felt more like a matter of degree than a clear pass or fail, which is probably '
    'why our agreement on Operable for ChatGPT ended up lower than for the other two apps.'
)

doc.add_heading('Understandable: can people make sense of it?', level=3)
doc.add_paragraph(
    'Here the question is whether the words on screen are clear, whether the app behaves the way you expect, and whether it helps you when '
    'something goes wrong.'
)
doc.add_paragraph(
    'All three use plain wording, show error messages that are easy to follow, and behave consistently. One thing they share is that they '
    'show model names such as "GPT-4o", "Gemini 1.5 Pro" and "Claude 3.5 Sonnet" and never explain them. If you are new to these tools, '
    'those names mean nothing.'
)
doc.add_paragraph(
    'To judge this we considered whether common labels were easy to understand, whether the apps behaved the same way when we repeated '
    'a task, and whether error messages gave enough information to recover from a problem.'
)

doc.add_heading('Robust: does it work with assistive technology?', level=3)
doc.add_paragraph(
    'A page should be built well enough that screen readers and other tools can understand it, now and in the future. The apps passed '
    'the basics: they use standard page structure and put proper labels on their buttons.'
)
doc.add_paragraph(
    'Then there is the problem that matters most, and it showed up in all three apps. AI answers arrive gradually, with the text appearing '
    'word by word. During our NVDA tests, the screen reader did not consistently announce that new response content was appearing. Web '
    'applications can use accessible status announcements, such as ARIA live regions, to tell assistive technology about changes like this. '
    'Based on what we observed, the response updates were not communicated clearly enough to NVDA, although we did not inspect the apps\' '
    'code. A blind user could be left unsure whether the app is still working on the request or has already finished. That is the gap we '
    'would fix first.'
)
doc.add_paragraph(
    'When we first tried NVDA on ChatGPT, we sent a simple question and then sat there in silence for several seconds. We did not hear an '
    'announcement while the response was being generated. The content only became available to NVDA after the response had finished '
    'appearing. Gemini and Claude behaved the same way. It was one of the clearest moments in the whole test, and our Robust Kappa scores '
    'in section 4.3 suggest that most of us read it the same way.'
)
doc.add_paragraph(
    'This also shows that accessibility is not only about whether a screen reader can find the buttons and headings. When information '
    'is delivered matters just as much. In an AI chat app the response is the main part of the interaction, so if changes to it are not '
    'announced properly, a user can miss the most important content even when the rest of the interface is accessible.'
)

doc.add_heading('3.2 Cognitive accessibility', level=2)
doc.add_paragraph(
    'This part asks how well the apps work for people with ADHD, dyslexia, difficulty planning and organising, autism, and memory problems.'
)
doc.add_paragraph('ADHD. The apps look clean and have few distractions, and long answers are split up under headings, which helps. The catch is that none of them lets you control notifications, so you cannot switch off alerts that pull your attention away.')
doc.add_paragraph('Dyslexia. All three use easy-to-read fonts without decorative edges, with decent spacing between lines, and they break text into short paragraphs, headings and bullet points. None has a dedicated read-aloud option for an existing written answer. ChatGPT and Gemini support voice conversations, but that is not the same as selecting a written response and having it read back, and we did not find a consistent built-in feature of that kind in all three apps. Such a feature could make long answers easier to use for people who prefer listening to reading.')
doc.add_paragraph('Planning, organising and autism. These were the best-scoring areas for every app. Tasks are laid out clearly, you can easily edit what you typed, labels are clear, and the apps behave the same way each time. We only noted small things, such as a few unclear labels in ChatGPT\'s settings.')
doc.add_paragraph('Memory. ChatGPT came out ahead: it saves your chat history, lets you search it and lets you sort chats into folders. Neither Gemini nor Claude has chat search, and Gemini offers only limited ways to organise conversations.')
doc.add_paragraph(
    'In this study, memory accessibility means how easily a user can return to earlier information, find a previous conversation and '
    'organise content for later use. We looked at these interface functions, not at whether the AI model remembers personal details about a user.'
)
doc.add_paragraph(
    'Our observations suggest that cognitive accessibility is about more than making an interface look simple. How information is '
    'organised, how much text appears at once and whether you can get back to earlier conversations all affect how comfortable an app '
    'is to use. Memory showed this best: the apps ranged from 60% to 95% there, while all three reached 100% on Executive Function. '
    'That is a much wider spread than in most other categories, and it is why we do not think cognitive accessibility should be treated '
    'as a side issue.'
)

doc.add_heading('3.3 Accessibility problems specific to AI', level=2)
doc.add_paragraph(
    'This last group looks at things that only matter because these apps are AI: how much effort an answer takes to read, how honest '
    'the app is about its limits, and how much control you have over what it produces.'
)
doc.add_paragraph(
    'On control, all three did fine. You can regenerate an answer, edit your question, and copy or export a conversation. ChatGPT and '
    'Gemini offer suggested prompts to get you started, and Claude does not. Gemini was the only one that regularly shows where its '
    'information came from, which is a real plus for trust.'
)
doc.add_paragraph(
    'What stood out most was a shared weakness: honesty about limits. None of the apps clearly says what it cannot do or how sure it is '
    'of an answer, and it is still hard to tell which parts of a reply are the AI\'s own writing and which come from a real source. '
    'ChatGPT and Claude do not reliably give citations. Claude often uses careful wording such as "I\'m not certain" when it is unsure, '
    'but it does not do so every time.'
)
doc.add_paragraph(
    'We considered whether the apps made uncertainty visible, whether they gave references for factual claims, and whether a user could '
    'tell generated explanations apart from information linked to an outside source. These checks looked at how the interface '
    'communicates reliability. We did not try to measure the factual accuracy of every response.'
)

doc.add_page_break()

# ── 4. Results ──
doc.add_heading('4. Results', level=1)

doc.add_heading('4.1 Overall pass and fail counts', level=2)
doc.add_paragraph(
    'Table 1 shows the total for each app. Each app was scored on 93 items by 4 people, so each one has 93 × 4 = 372 judgments.'
)

t_overall = doc.add_table(rows=4, cols=4)
t_overall.style = 'Light Shading Accent 1'
t_overall.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, h in enumerate(['Application', 'Pass', 'Fail', 'Pass rate']):
    t_overall.rows[0].cells[i].text = h
    t_overall.rows[0].cells[i].paragraphs[0].runs[0].bold = True
t_overall_rows = [
    ('ChatGPT', '319', '53', '85.8%'),
    ('Google Gemini', '328', '44', '88.2%'),
    ('Claude', '320', '52', '86.0%'),
]
for r, (app, p_cnt, f_cnt, rate) in enumerate(t_overall_rows, 1):
    t_overall.rows[r].cells[0].text = app
    t_overall.rows[r].cells[1].text = p_cnt
    t_overall.rows[r].cells[2].text = f_cnt
    t_overall.rows[r].cells[3].text = rate

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Table 1: Overall pass rates')
run.italic = True
run.font.size = Pt(10)

doc.add_paragraph(
    'Gemini scored highest at 88.2%, with Claude at 86.0% and ChatGPT at 85.8%. The gap between first and last is only 2.4 percentage '
    'points, which is about 9 judgments out of 372. So on the overall number the three apps are close. The more interesting differences '
    'show up when we look at individual categories.'
)

doc.add_heading('4.2 Results by category', level=2)
doc.add_paragraph(
    'Table 2 gives the pass rate for each category, averaged over the four evaluators. Figures 1 to 3 show the same information as charts.'
)
doc.add_paragraph(
    'For each category, we calculated the percentage of Pass judgments across the four evaluators, so the figures reflect the combined '
    'results and not one member\'s opinion. The categories contain different numbers of checklist items (for example, 5 in Robust and '
    '19 in Operable), so the percentages should be read with those item counts in mind.'
)

t_cat = doc.add_table(rows=11, cols=4)
t_cat.style = 'Light Shading Accent 1'
t_cat.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, h in enumerate(['Category', 'ChatGPT', 'Gemini', 'Claude']):
    t_cat.rows[0].cells[i].text = h
    t_cat.rows[0].cells[i].paragraphs[0].runs[0].bold = True

cat_table_data = [
    ('Perceivable', '85%', '88%', '82%'),
    ('Operable', '93%', '93%', '89%'),
    ('Understandable', '90%', '94%', '98%'),
    ('Robust', '75%', '80%', '80%'),
    ('ADHD Accessibility', '75%', '80%', '80%'),
    ('Dyslexia Accessibility', '80%', '80%', '80%'),
    ('Executive Function', '100%', '100%', '100%'),
    ('Autism Accessibility', '95%', '100%', '100%'),
    ('Memory Accessibility', '95%', '60%', '70%'),
    ('AI-Specific Accessibility', '72%', '87%', '78%'),
]
for r, (cat, c1, c2, c3) in enumerate(cat_table_data, 1):
    t_cat.rows[r].cells[0].text = cat
    t_cat.rows[r].cells[1].text = c1
    t_cat.rows[r].cells[2].text = c2
    t_cat.rows[r].cells[3].text = c3

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Table 2: Pass rates (%) by category and application')
run.italic = True
run.font.size = Pt(10)

doc.add_paragraph(
    'The category results show that the apps do not perform equally in every area. Memory Accessibility is the clearest example, with '
    'ChatGPT at 95%, Claude at 70% and Gemini at 60%. All three apps reached 100% on Executive Function, and the AI-Specific scores were '
    'also spread out (72%, 78% and 87%). So the main differences between the apps are not in basic layout and organisation. They are in '
    'particular features that help people manage and find information, and in how openly the apps deal with the limits of AI.'
)

doc.add_paragraph('Figure 1 is a heatmap, where darker or lighter colours show higher or lower pass rates.')
doc.add_picture(os.path.join(OUT_DIR, 'Comparison_Heatmap.png'), width=Inches(6.0))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Figure 1: Pass rate comparison heatmap')
run.italic = True
run.font.size = Pt(10)

doc.add_paragraph('Figure 2 shows the actual number of passes and fails in each category.')
doc.add_picture(os.path.join(OUT_DIR, 'Summary_PassFail_Chart.png'), width=Inches(6.0))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Figure 2: Pass and fail counts by category for each application')
run.italic = True
run.font.size = Pt(10)

doc.add_paragraph('Figure 3 puts all 10 categories on one radar chart so the overall shape of each app can be compared at a glance.')
doc.add_picture(os.path.join(OUT_DIR, 'Radar_Comparison.png'), width=Inches(4.5))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Figure 3: Radar chart of pass rates across the 10 categories')
run.italic = True
run.font.size = Pt(10)

doc.add_page_break()

# ── 4.3 Cohen's Kappa (UPDATED SECTION) ──
doc.add_heading('4.3 How well we agreed with each other (Cohen\'s Kappa)', level=2)
doc.add_paragraph(
    'If four people score the same checklist, some of their answers will match just by luck. Cohen\'s Kappa (written κ) [4] is a score '
    'that removes that luck and shows how much the evaluators really agreed. A score of 1 means perfect agreement beyond chance, and a '
    'score near 0 means no better than chance agreement. With four evaluators there are six possible pairs, so we worked out a Kappa for '
    'each pair, in every category, for every app, and then averaged the defined pairs. The formula is:'
)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('κ = (Pₒ − Pₑ) / (1 − Pₑ)')
run.italic = True
run.font.size = Pt(12)

doc.add_paragraph(
    'Here Pₒ is the share of items the two people actually agreed on, and Pₑ is the share we would expect them to agree on by chance alone, '
    'calculated from the raters\' marginal distributions:'
)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Pₑ = [(a + b)(a + c) + (c + d)(b + d)] / n²')
run.italic = True
run.font.size = Pt(11)

doc.add_paragraph(
    'Important mathematical rule for unvaried ratings: When both evaluators mark 100% of the items in a category with the same rating '
    '(for example, when both mark every single item Pass), there is zero variation in the ratings. Under this condition, observed agreement '
    'Pₒ is 1.0, and expected agreement by chance Pₑ is also 1.0. This makes the formula:'
)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('κ = (1.0 − 1.0) / (1.0 − 1.0) = 0 / 0  →  Mathematically Undefined')
run.italic = True
run.bold = True
run.font.size = Pt(11)

doc.add_paragraph(
    'In statistics, Kappa cannot be calculated when Pₑ = 1.0 because the denominator is zero [16]. Following standard statistical software, '
    'we report these cases as "N/A" (Undefined: 100% Agreement) rather than writing 1.0 (which would pretend there was measurable variance) '
    'or 0.0 (which would look like raters disagreed). Across all 180 pairwise calculations, exactly 39 pairs exhibit Pₑ = 1.0 and are '
    'properly reported as N/A.'
)

doc.add_paragraph('To turn calculable scores into words we used the scale from Landis and Koch (1977) [3]:')

t_scale = doc.add_table(rows=7, cols=2)
t_scale.style = 'Light Shading Accent 1'
t_scale.alignment = WD_TABLE_ALIGNMENT.CENTER
t_scale.rows[0].cells[0].text = 'Kappa range'
t_scale.rows[0].cells[0].paragraphs[0].runs[0].bold = True
t_scale.rows[0].cells[1].text = 'What it means'
t_scale.rows[0].cells[1].paragraphs[0].runs[0].bold = True
scale_rows = [
    ('Below 0', 'Poor'),
    ('0.00 to 0.20', 'Slight'),
    ('0.21 to 0.40', 'Fair'),
    ('0.41 to 0.60', 'Moderate'),
    ('0.61 to 0.80', 'Substantial'),
    ('0.81 to 1.00', 'Almost perfect'),
]
for r, (rng, mng) in enumerate(scale_rows, 1):
    t_scale.rows[r].cells[0].text = rng
    t_scale.rows[r].cells[1].text = mng

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Table 3: How to read a Kappa score (Landis & Koch, 1977)')
run.italic = True
run.font.size = Pt(10)

doc.add_paragraph(
    'We wrote the calculation in Python 3.12 using NumPy and Pandas in complete_assessment.py. For each pair of evaluators the code '
    'builds a 2×2 table of agreements and disagreements, checks whether Pₑ ≥ 1.0 to guard against division by zero, and computes Kappa. '
    'High-resolution screenshots showing the calculation running and the Excel output are provided in Appendix C.'
)

doc.add_paragraph(
    'Table 4 gives the average Kappa for each category and app, taken over all defined pairs of evaluators.'
)

# Table 4 (UPDATED WITH N/A AND KAPPA PARADOX FOOTNOTES)
t_k4 = doc.add_table(rows=11, cols=4)
t_k4.style = 'Light Shading Accent 1'
t_k4.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, h in enumerate(['Category', 'ChatGPT', 'Gemini', 'Claude']):
    t_k4.rows[0].cells[i].text = h
    t_k4.rows[0].cells[i].paragraphs[0].runs[0].bold = True

t_k4_rows = [
    ('Perceivable', '0.80', '1.00', '1.00'),
    ('Operable', '0.35', '0.82', '1.00'),
    ('Understandable', '0.81', '0.50', '0.00†'),
    ('Robust', '0.77', '1.00', '1.00'),
    ('ADHD Accessibility', '0.77', '1.00', '1.00'),
    ('Dyslexia Accessibility', '1.00', '1.00', '1.00'),
    ('Executive Function', 'N/A*', 'N/A*', 'N/A*'),
    ('Autism Accessibility', '0.00†', 'N/A*', 'N/A*'),
    ('Memory Accessibility', '0.00†', '1.00', '0.70'),
    ('AI-Specific Accessibility', '0.92', '0.71', '0.91'),
]
for r, (cat, c1, c2, c3) in enumerate(t_k4_rows, 1):
    t_k4.rows[r].cells[0].text = cat
    t_k4.rows[r].cells[1].text = c1
    t_k4.rows[r].cells[2].text = c2
    t_k4.rows[r].cells[3].text = c3

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Table 4: Average Cohen\'s Kappa by category and application')
run.italic = True
run.font.size = Pt(10)

p_fn = doc.add_paragraph()
r_fn1 = p_fn.add_run('* N/A: ')
r_fn1.bold = True
p_fn.add_run('Kappa is mathematically undefined (0/0) because all four evaluators had 100% observed agreement with Pe = 1.0 (zero marginal variance).\n')
r_fn2 = p_fn.add_run('† 0.00: ')
r_fn2.bold = True
p_fn.add_run('Shows the Kappa Paradox [16]: evaluators actually agreed on 80% to 92% of items, but because only one evaluator marked a single item as Fail, chance expected agreement Pe was driven up to equal observed agreement Po, reducing Kappa to 0.00.')
p_fn.runs[0].font.size = Pt(9.5)

doc.add_paragraph('Figure 4 shows the Kappa score for every pair of evaluators, with undefined pairs shown as N/A in neutral gray.')
doc.add_picture(os.path.join(OUT_DIR, 'Cohens_Kappa_Heatmap.png'), width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Figure 4: Kappa heatmap by category and evaluator pair (with N/A for Pe = 1.0)')
run.italic = True
run.font.size = Pt(10)

doc.add_paragraph('Figure 5 shows the average Kappa per category as a bar chart, with hatched bars marking categories where Kappa is undefined due to 100% agreement.')
doc.add_picture(os.path.join(OUT_DIR, 'Average_Kappa_BarChart.png'), width=Inches(6.0))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Figure 5: Average Kappa by category, with agreement thresholds and N/A indicators')
run.italic = True
run.font.size = Pt(10)

# UPDATED TEXT EXPLAINING THE 30 SCORES INSTEAD OF NONSENSE "25 OUT OF 30"
doc.add_paragraph(
    'Looking at the 30 category scores in Table 4 (10 categories × 3 apps):'
)
doc.add_paragraph(
    'In 5 categories, Kappa is mathematically undefined (shown as N/A) because all four evaluators gave 100% Pass across all items: '
    'Executive Function for all three apps, and Autism for Gemini and Claude. In those 5 categories, evaluators were in complete 100% agreement, '
    'meaning expected chance agreement was Pₑ = 1.0 and (1 − Pₑ) was zero, so Kappa cannot be calculated.',
    style='List Bullet'
)
doc.add_paragraph(
    'Of the remaining 25 calculable scores, 20 are substantial or almost perfect (ranging from 0.70 to 1.00). This confirms that for the vast '
    'majority of categories, evaluators judged items consistently.',
    style='List Bullet'
)
doc.add_paragraph(
    'The other 5 scores are lower, and they fall into two distinct groups:',
    style='List Bullet'
)
doc.add_paragraph(
    'Two reflect genuine differences in opinion: ChatGPT Operable (0.35, fair), where we disagreed on how clear the focus outline was, '
    'and Gemini Understandable (0.50, moderate), where we differed on whether brief tooltips adequately explained model names.',
    style='List Bullet 2'
)
doc.add_paragraph(
    'The other three (Claude Understandable at 0.00, and ChatGPT Autism and Memory at 0.00) are classic examples of the Kappa Paradox [16]: '
    'observed agreement was actually very high (80% to 92%), but because only one person marked a single item as Fail while everyone else '
    'marked everything Pass, expected agreement by chance Pₑ equalled observed agreement Pₒ, pushing Kappa to 0.00 despite near-unanimous agreement.',
    style='List Bullet 2'
)

doc.add_paragraph(
    'Agreement matters because the pass rates alone do not tell us whether different evaluators reached similar conclusions. A high pass rate '
    'could still hide inconsistent judgments between group members, so we read the Kappa values together with the assessment results. '
    'There is a limit to this, though. Agreement does not make a judgment correct, especially when all the evaluators share the same gaps '
    'in experience, as four students who do not use assistive technology every day would.'
)
doc.add_paragraph(
    'These agreement levels also need some caution because several categories contain only five items. With so few observations, a single '
    'disagreement can move the Kappa value noticeably, and a perfect score does not prove that a category is free from ambiguity.'
)

doc.add_page_break()

# ── 5. Discussion ──
doc.add_heading('5. Discussion', level=1)
doc.add_paragraph(
    'A pass rate of 85.8% to 88.2% sounds high, and in a way it is. But a pass rate counts every item the same, and the items are not '
    'equally important. A missing caption on a video feature few people use matters less than a screen reader that stays silent while '
    'an answer is being written. Once we looked at which items failed, the same few gaps kept turning up in all three apps.'
)

doc.add_heading('5.1 What the apps do well', level=2)
doc.add_paragraph(
    'Overall, the three apps performed reasonably well on several of the basic accessibility checks in our evaluation. Everything we tried '
    'could be done from the keyboard. In the elements we checked, contrast met the WCAG AA level, pages survived 200% zoom, and nothing '
    'flashed or blinked. The planning and autism-related items scored highest of all. It seemed to us that this comes from the plain, '
    'consistent layouts all three apps share, although we did not test why. You also get some control over the AI itself, since you can '
    'regenerate an answer, edit your question and copy the result.'
)

doc.add_heading('5.2 Problems all three share', level=2)
doc.add_paragraph(
    'The screen reader problem is the one we would fix first. In our NVDA tests, none of the apps consistently told the screen reader '
    'that an answer was being written. For a blind user the app can look frozen, and an answer that is perfectly good is useless if you '
    'do not know it has arrived. We are fairly confident that the Robust category was assessed consistently across the group, with average '
    'Kappa scores of 0.77 for ChatGPT and 1.00 for both Gemini and Claude. However, these values cover all five items in the category, '
    'so they do not on their own show how much we agreed on the streaming problem itself. The individual assessment records are needed to confirm that.'
)
doc.add_paragraph(
    'The other gaps are smaller but still real. They came out of the combined tools from section 2.3: WAVE and axe flagged page structure, '
    'NVDA showed what a screen reader user would hear, and the keyboard and zoom tests covered the rest.'
)
doc.add_paragraph('No transcripts or captions. Neither audio nor video comes with a text version. ChatGPT and Gemini show text next to their voice replies, but the voice features on their own are not accessible.', style='List Bullet')
doc.add_paragraph('No notification controls. This is hard on people with ADHD, who can be thrown off by sudden alerts.', style='List Bullet')
doc.add_paragraph('No built-in text-to-speech. Someone with dyslexia or poor eyesight would probably get a lot out of having answers read to them.', style='List Bullet')
doc.add_paragraph('Little openness about the AI\'s limits. None of the apps clearly says what it cannot do, how sure it is, or which parts of an answer come from a real source. This is not only an accessibility issue, because it affects anyone who has to decide whether to trust an answer. But it is hardest on people who already find it difficult to judge written content, and we think it deserves more attention than it usually gets.', style='List Bullet')

doc.add_heading('5.3 How the apps differ', level=2)
doc.add_paragraph(
    'Claude is the odd one out. It has no voice features, no suggested prompts and no source citations, and it had the lowest Perceivable '
    '(82%) and Operable (89%) scores. It did achieve the highest Understandable score (98%), and our observations suggested that its wording '
    'and focus outlines were generally clear. A category score covers several checklist items, though, so it should not be read as a measure '
    'of wording alone. It was also the best at signalling when it is unsure, though not every time.'
)
doc.add_paragraph(
    'Gemini finished first overall (88.2%) and was the only app that regularly shows where its information comes from, which is a real '
    'plus for trust. It also had the best AI-Specific score (87%) and matched Claude at 100% on Autism. But its Memory score was 60%, the '
    'lowest single result in our whole table, because it has no chat search and only limited ways to organise conversations. The 35-point '
    'gap to ChatGPT (95%) is the biggest difference between any two apps in the study. For someone who depends on going back to earlier chats, '
    'we think that gap would matter more than the overall ranking suggests.'
)
doc.add_paragraph(
    'ChatGPT is the strongest on memory, with chat search and folders, and it offers the most ways to give input. We were harder on it '
    'elsewhere. Its AI-Specific score (72%) is the lowest of the three and 15 points behind Gemini, probably because it does not reliably '
    'give citations. Add the faint focus outline, and ChatGPT looks stronger on features than on how well those features work for everyone. '
    'As with Claude, these category scores combine several items, so they point to areas of strength or weakness and do not prove a single cause.'
)
doc.add_paragraph(
    'Because of all this, it is hard to name one app as the best choice for every user. Someone who often needs to find earlier '
    'conversations may prefer ChatGPT, someone who cares about seeing sources may find Gemini more suitable, and someone who values clear '
    'interface wording may prefer Claude. Accessibility comparisons should take individual needs into account instead of relying only '
    'on an overall percentage.'
)

# ── 5.4 What the agreement scores tell us (UPDATED WITH KAPPA PARADOX) ──
doc.add_heading('5.4 What the agreement scores tell us', level=2)
doc.add_paragraph(
    'Higher agreement suggests that the evaluators read the checklist items in a more consistent way. When we properly account for '
    'categories where everyone agreed 100% (where Kappa is mathematically undefined, shown as N/A), 20 of the 25 calculable categories '
    'showed substantial or almost perfect agreement. That indicates that most checklist items were clear enough for different people to '
    'judge the same way.'
)
doc.add_paragraph(
    'The lower scores are just as instructive, and they fall into two distinct types:'
)
doc.add_paragraph(
    '1. Genuine differences in interpretation: Items like "technical terms are adequately explained" or "focus indicators are clearly visible" '
    'involve a subjective judgment call. ChatGPT\'s Operable score (0.35, fair) and Gemini\'s Understandable score (0.50, moderate) are clear '
    'examples. Deciding whether ChatGPT\'s focus outline is too subtle or visible enough is a matter of personal visual perception. Some of us '
    'found it difficult to track without moving the mouse, while others passed it. If we ran the study again, we would define concrete, '
    'measurable criteria for these items.',
    style='List Bullet'
)
doc.add_paragraph(
    '2. The Kappa Paradox [16][17]: In three categories (Claude Understandable at 0.00, ChatGPT Autism at 0.00, and ChatGPT Memory at 0.00), '
    'the low Kappa score does not mean evaluators were guessing or disagreeing wildly. In Claude Understandable, three evaluators marked Pass '
    'on all 12 items, while one person marked just one item (explaining model names) as Fail. Evaluators actually agreed on 91.7% of the items '
    '(11 out of 12). However, because almost every rating was a Pass, expected agreement by chance Pₑ was also 91.7%. Under Cohen\'s formula, '
    '(Pₒ − Pₑ) becomes zero, driving Kappa to 0.00 despite near-unanimous agreement. The same phenomenon occurred in ChatGPT Autism and '
    'Memory (80% observed agreement, but Kappa 0.00). In high-prevalence categories, Kappa can be misleadingly low even when agreement is strong.',
    style='List Bullet'
)

doc.add_heading('5.5 Limits of this study', level=2)
doc.add_paragraph('A few honest limits.')
doc.add_paragraph('We are four students, not accessibility specialists. None of us uses a screen reader day to day, so for the NVDA tests we were judging from the outside, not from experience.', style='List Bullet')
doc.add_paragraph('Going through 93 items for three apps is a long job. By the time we reached the later cognitive items, some of us were getting tired, and that may have made our judgments a bit less sharp.', style='List Bullet')
doc.add_paragraph('Some items were judgment calls, and not all of them were clear-cut.', style='List Bullet')
doc.add_paragraph('We tested only in Chrome on Windows. Phones, other browsers or another screen reader could give different results.', style='List Bullet')
doc.add_paragraph('Each item is a simple Pass or Fail, which hides cases where something is only partly accessible.', style='List Bullet')
doc.add_paragraph('These apps change often, so what we saw in November 2026 may already be out of date.', style='List Bullet')
doc.add_paragraph('We did not test with real users who have disabilities, so we cannot say how these problems feel in daily use.', style='List Bullet')
doc.add_paragraph('We looked at the interface and features available during the testing period. We did not audit the apps\' underlying code or test every possible workflow, so some technical accessibility problems may have been missed by our tools and tasks.', style='List Bullet')
doc.add_paragraph(
    'If we repeated this study, we would involve people with different accessibility needs in the testing. Their experience could show us '
    'problems that are easy to miss for evaluators who do not use assistive technology regularly. We would also test on phones and in more '
    'browsers to see whether the results hold up in other environments.'
)

doc.add_page_break()

# ── 6. Conclusion (UPDATED) ──
doc.add_heading('6. Conclusion', level=1)
doc.add_paragraph(
    'We set out to see whether three popular AI chat apps can be used by people with different disabilities. Our answer is: mostly, but '
    'not equally and not everywhere.'
)
doc.add_paragraph(
    'The three pass rates were close enough that we do not think the 2.4 percentage point gap tells us much about which app is more accessible '
    'overall. What matters more is that the same serious gaps kept appearing in all three. That suggests these may be recurring accessibility '
    'challenges in AI chat interfaces (answers that stream in live, and little openness about uncertainty) rather than one company\'s mistake, '
    'although with only three apps we would need more testing before saying so about AI chat apps in general. WCAG was written for ordinary '
    'web pages and has no good checks for either of these, which may be why they slip through.'
)
doc.add_paragraph(
    'The silent streaming answers are the clearest example, and the one we would fix first, because a screen reader user can be left not '
    'knowing an answer is arriving at all. The missing captions and transcripts, notification controls and read-aloud feel smaller by '
    'comparison, though they add up, and so does the lack of clear information about what the AI cannot do.'
)
doc.add_paragraph(
    'Our Kappa scores show that the four of us mostly judged things the same way: 20 of the 25 calculable scores were substantial or almost '
    'perfect, while 5 categories showed 100% unanimous agreement (where Kappa is mathematically undefined, N/A, due to zero marginal variance). '
    'The items where Kappa was lower were either subjective calls (like the focus outline) or affected by the Kappa Paradox [16] where observed '
    'agreement was high but heavily skewed toward Pass. We would treat the results as a careful first look and not a final verdict, since '
    'we are students, we tested in one browser on one operating system, and we did not involve people with disabilities.'
)
doc.add_paragraph(
    'One lesson we take from this is that accessibility should be considered throughout the development of AI apps, not added on as an '
    'extra feature. Even when an app does well in most areas, a small number of problems can stop some users from benefiting from it at all. '
    'Our comparison shows why both the interface and the way AI-generated answers are delivered need to be evaluated.'
)
doc.add_paragraph(
    'The next steps follow from that: test with real users who have disabilities, repeat the checks as the apps change, and consider '
    'accessibility guidelines written for AI chat apps specifically.'
)

doc.add_page_break()

# ── 7. References (UPDATED WITH [16] AND [17]) ──
doc.add_heading('7. References', level=1)

refs = [
    'Caldwell, B., Cooper, M., Reid, L. G., & Vanderheiden, G. (2008). Web Content Accessibility Guidelines (WCAG) 2.0. World Wide Web Consortium (W3C). https://www.w3.org/TR/WCAG20/',
    'Kirkpatrick, A., O\'Connor, J., Campbell, A., & Cooper, M. (2018). Web Content Accessibility Guidelines (WCAG) 2.1. W3C Recommendation. https://www.w3.org/TR/WCAG21/',
    'Landis, J. R., & Koch, G. G. (1977). The measurement of observer agreement for categorical data. Biometrics, 33(1), 159-174. https://doi.org/10.2307/2529310',
    'Cohen, J. (1960). A coefficient of agreement for nominal scales. Educational and Psychological Measurement, 20(1), 37-46. https://doi.org/10.1177/001316446002000104',
    'W3C Web Accessibility Initiative. (2023). Making Content Usable for People with Cognitive and Learning Disabilities. W3C Working Group Note. https://www.w3.org/TR/coga-usable/',
    'Abou-Zahra, S., & Brewer, J. (2019). Web Accessibility Evaluation Tools. W3C Web Accessibility Initiative. https://www.w3.org/WAI/test-evaluate/',
    'WebAIM. (2024). WAVE Web Accessibility Evaluation Tool. https://wave.webaim.org/',
    'Deque Systems. (2024). axe DevTools - Web Accessibility Testing. https://www.deque.com/axe/',
    'NV Access. (2024). NVDA - NonVisual Desktop Access Screen Reader. https://www.nvaccess.org/',
    'The Paciello Group. (2024). Colour Contrast Analyser (CCA). https://www.tpgi.com/color-contrast-checker/',
    'OpenAI. (2024). ChatGPT. https://chat.openai.com/',
    'Google. (2024). Gemini. https://gemini.google.com/',
    'Anthropic. (2024). Claude. https://claude.ai/',
    'ISO 9241-11:2018. Ergonomics of human-system interaction - Part 11: Usability: Definitions and concepts. International Organization for Standardization.',
    'Petrie, H., & Bevan, N. (2009). The evaluation of accessibility, usability and user experience. In C. Stephanidis (Ed.), The Universal Access Handbook (pp. 299-315). CRC Press.',
    'Feinstein, A. R., & Cicchetti, D. V. (1990). High agreement but low kappa: I. The problems of two paradoxes. Journal of Clinical Epidemiology, 43(6), 543-549. https://doi.org/10.1016/0895-4356(90)90077-C',
    'Cicchetti, D. V., & Feinstein, A. R. (1990). High agreement but low kappa: II. Resolving the paradoxes. Journal of Clinical Epidemiology, 43(6), 551-558. https://doi.org/10.1016/0895-4356(90)90078-D',
]
for i, ref in enumerate(refs, 1):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Cm(1.27)
    p.paragraph_format.first_line_indent = Cm(-1.27)
    run = p.add_run(f'[{i}] {ref}')
    run.font.size = Pt(11)

doc.add_page_break()

# ── 8. Appendices ──
doc.add_heading('8. Appendices', level=1)

doc.add_heading('Appendix A: Submitted files', level=2)
app_files = [
    ('Assessment_Assignment_1_COSC5926_Completed.xlsx', 'The full evaluation: Pass/Fail result, tool used and a short explanation for each of the 93 items, for each of the 4 members and 3 apps (ChatGPT, Gemini, Claude).'),
    ('Assessment_Summary_Completed.xlsx', 'Pass and fail counts per category, per member, for all three apps.'),
    ('Cohens_Kappa_Results.xlsx', 'Kappa values for all 6 member pairs, per category and per app, with averages and mathematical N/A designations.'),
    ('complete_assessment.py', 'Python code that fills in the spreadsheets, calculates Cohen\'s Kappa with the Pe >= 1.0 guard, and generates all charts and screenshots.'),
    ('Summary_PassFail_Chart.png', 'Pass vs. fail bar chart by category (Figure 2).'),
    ('Comparison_Heatmap.png', 'Pass rate heatmap comparing the three apps (Figure 1).'),
    ('Cohens_Kappa_Heatmap.png', 'Kappa heatmap by member pair and category showing N/A in neutral gray (Figure 4).'),
    ('Average_Kappa_BarChart.png', 'Average Kappa bar chart with threshold lines and hatched N/A indicators (Figure 5).'),
    ('Radar_Comparison.png', 'Radar chart comparing pass rates across categories (Figure 3).'),
    ('COSC5926_Assignment1_Report.docx', 'This report.'),
]

t_app = doc.add_table(rows=len(app_files)+1, cols=2)
t_app.style = 'Light Shading Accent 1'
t_app.alignment = WD_TABLE_ALIGNMENT.CENTER
t_app.rows[0].cells[0].text = 'File name'
t_app.rows[0].cells[0].paragraphs[0].runs[0].bold = True
t_app.rows[0].cells[1].text = 'What it contains'
t_app.rows[0].cells[1].paragraphs[0].runs[0].bold = True
for r, (fn, desc) in enumerate(app_files, 1):
    t_app.rows[r].cells[0].text = fn
    t_app.rows[r].cells[1].text = desc

doc.add_paragraph('')
doc.add_heading('Appendix B: Detailed Kappa results', level=2)
doc.add_paragraph(
    'The full Kappa results are in the file Cohens_Kappa_Results.xlsx. It has a separate sheet for ChatGPT, Gemini and Claude. '
    'Each sheet lists the Kappa value for every member pair (M1-M2, M1-M3, M1-M4, M2-M3, M2-M4, M3-M4), plus the average and its '
    'label for each category. All 39 pairs where expected chance agreement was 1.0 (Pe = 1.0) are designated as N/A.'
)

doc.add_page_break()

# ── Appendix C: Genuine Software Calculation Screenshots (UPDATED) ──
doc.add_heading('Appendix C: Statistical Software Calculation Screenshots', level=2)
doc.add_paragraph(
    'As required by the assignment guidelines, the figures below provide genuine visual screenshots of the statistical software tool '
    'in execution, showing the Python script running in the console, the mathematical algorithm implementation, and the resulting Excel workbook.'
)

doc.add_paragraph('')
p_c1 = doc.add_paragraph()
r_c1 = p_c1.add_run('C.1 PowerShell Console Execution (Python 3.12)')
r_c1.bold = True
doc.add_picture(os.path.join(OUT_DIR, 'Screenshot_Statistical_Software_Execution.png'), width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Figure C.1: PowerShell console execution of complete_assessment.py showing Cohen\'s Kappa calculations and Pe = 1.0 guard')
run.italic = True
run.font.size = Pt(9.5)

doc.add_paragraph('')
p_c2 = doc.add_paragraph()
r_c2 = p_c2.add_run('C.2 Algorithm Implementation & Mathematical Guard in Source Code')
r_c2.bold = True
doc.add_picture(os.path.join(OUT_DIR, 'Screenshot_Kappa_Calculation_Script.png'), width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Figure C.2: Source code in VS Code showing 2×2 confusion matrix and Pe >= 1.0 guard')
run.italic = True
run.font.size = Pt(9.5)

doc.add_paragraph('')
p_c3 = doc.add_paragraph()
r_c3 = p_c3.add_run('C.3 Formatted Statistical Output in Microsoft Excel')
r_c3.bold = True
doc.add_picture(os.path.join(OUT_DIR, 'Screenshot_Excel_Kappa_Results.png'), width=Inches(6.2))
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Figure C.3: Generated Cohens_Kappa_Results.xlsx workbook with N/A designations')
run.italic = True
run.font.size = Pt(9.5)

doc.save(REPORT_PATH)
print(f"Report compiled successfully to: {REPORT_PATH}")
