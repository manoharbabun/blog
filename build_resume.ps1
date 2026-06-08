$ErrorActionPreference = 'Stop'

$outPath = 'c:\Users\ManoharPattanayak\OneDrive - CXIO Technologies Pvt Ltd\Desktop\blog\Manohar_Pattanayak_Resume_ATS.docx'

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()

# Page margins (compact but ATS-safe)
$doc.PageSetup.TopMargin    = 30
$doc.PageSetup.BottomMargin = 30
$doc.PageSetup.LeftMargin   = 45
$doc.PageSetup.RightMargin  = 45

$sel = $word.Selection
$sel.Font.Name = 'Calibri'

# wd constants
$wdAlignLeft   = 0
$wdAlignCenter = 1

function Add-Line {
    param(
        [string]$Text,
        [int]$Size = 11,
        [bool]$Bold = $false,
        [bool]$Italic = $false,
        [int]$Align = 0,
        [int]$SpaceAfter = 4,
        [int]$SpaceBefore = 0,
        [bool]$AllCaps = $false
    )
    $s = $word.Selection
    $s.ParagraphFormat.Alignment   = $Align
    $s.ParagraphFormat.SpaceAfter  = $SpaceAfter
    $s.ParagraphFormat.SpaceBefore = $SpaceBefore
    $s.ParagraphFormat.LineSpacingRule = 0  # single
    $s.Font.Name = 'Calibri'
    $s.Font.Size = $Size
    $s.Font.Bold = [int]$Bold
    $s.Font.Italic = [int]$Italic
    $s.Font.AllCaps = [int]$AllCaps
    $s.TypeText($Text)
    $s.TypeParagraph()
}

function Add-Heading {
    param([string]$Text)
    $s = $word.Selection
    $s.ParagraphFormat.Alignment   = 0
    $s.ParagraphFormat.SpaceBefore = 3
    $s.ParagraphFormat.SpaceAfter  = 1
    $s.ParagraphFormat.LineSpacingRule = 0
    $s.Font.Name = 'Calibri'
    $s.Font.Size = 11.5
    $s.Font.Bold = 1
    $s.Font.AllCaps = 1
    $s.Font.Color = 0  # black
    $s.TypeText($Text)
    $s.TypeParagraph()
    # bottom border on the heading paragraph
    $para = $doc.Paragraphs($doc.Paragraphs.Count - 1)
    $border = $para.Borders(-3)  # wdBorderBottom
    $border.LineStyle = 1
    $border.LineWidth = 8  # 1pt
    $border.Color = 0
    $s.Font.AllCaps = 0
}

function Add-Bullet {
    param([string]$Text)
    $s = $word.Selection
    $s.ParagraphFormat.Alignment   = 3  # justify
    $s.ParagraphFormat.SpaceAfter  = 1
    $s.ParagraphFormat.SpaceBefore = 0
    $s.ParagraphFormat.LineSpacingRule = 4  # wdLineSpaceExactly
    $s.ParagraphFormat.LineSpacing = 11.5
    $s.ParagraphFormat.LeftIndent  = 14
    $s.ParagraphFormat.FirstLineIndent = -10
    $s.Font.Name = 'Calibri'
    $s.Font.Size = 10
    $s.Font.Bold = 0
    $s.Font.Italic = 0
    $s.TypeText([char]0x2022 + '  ' + $Text)
    $s.TypeParagraph()
    # reset indent
    $s.ParagraphFormat.LeftIndent  = 0
    $s.ParagraphFormat.FirstLineIndent = 0
}

# Add a role line: Title (bold) left, dates right via tab stop
function Add-Role {
    param([string]$Left, [string]$Right)
    $s = $word.Selection
    $s.ParagraphFormat.Alignment   = 0
    $s.ParagraphFormat.SpaceBefore = 3
    $s.ParagraphFormat.SpaceAfter  = 0
    $s.ParagraphFormat.LineSpacingRule = 0
    $s.ParagraphFormat.TabStops.ClearAll()
    # right-aligned tab at right margin (~7.0 in usable width)
    $rightPos = $doc.PageSetup.PageWidth - $doc.PageSetup.LeftMargin - $doc.PageSetup.RightMargin
    [void]$s.ParagraphFormat.TabStops.Add($rightPos, 2)  # 2 = wdAlignTabRight
    $s.Font.Name = 'Calibri'
    $s.Font.Size = 11
    $s.Font.Bold = 1
    $s.TypeText($Left)
    $s.Font.Bold = 0
    $s.Font.Size = 10.5
    $s.TypeText([char]9 + $Right)
    $s.TypeParagraph()
    $s.ParagraphFormat.TabStops.ClearAll()
}

function Add-SubLine {
    param([string]$Text)
    $s = $word.Selection
    $s.ParagraphFormat.Alignment   = 0
    $s.ParagraphFormat.SpaceBefore = 0
    $s.ParagraphFormat.SpaceAfter  = 2
    $s.ParagraphFormat.LineSpacingRule = 0
    $s.Font.Name = 'Calibri'
    $s.Font.Size = 10
    $s.Font.Bold = 0
    $s.Font.Italic = 1
    $s.TypeText($Text)
    $s.TypeParagraph()
    $s.Font.Italic = 0
}

# ---------------- HEADER ----------------
Add-Line -Text 'MANOHAR PATTANAYAK' -Size 19 -Bold $true -Align $wdAlignCenter -SpaceAfter 1
Add-Line -Text 'AI Solution Security Architect' -Size 12 -Bold $false -Align $wdAlignCenter -SpaceAfter 2
Add-Line -Text 'Mumbai, India  |  +91 6372761005  |  manoharpattanayak3@gmail.com' -Size 10.5 -Align $wdAlignCenter -SpaceAfter 2
Add-Line -Text 'LinkedIn: linkedin.com/in/manohar-pattanayak  |  GitHub: github.com/manoharbabun' -Size 10.5 -Align $wdAlignCenter -SpaceAfter 6

# ---------------- SUMMARY ----------------
Add-Heading 'Professional Summary'
Add-Line -Text ('AI Solution Security Architect with nearly 3 years of experience designing secure architectures for AI-integrated, ' +
    'cloud-native, and enterprise applications across BFSI, Fintech, Insurance, and Healthcare. Specializes in AI security ' +
    'architecture, threat modeling, application security testing, and cloud security across AWS, Azure, and GCP, with a proven ' +
    'record of reducing risk exposure and achieving 100% audit and remediation compliance.') -Size 10 -Align 3 -SpaceAfter 3

# ---------------- CORE COMPETENCIES ----------------
Add-Heading 'Core Competencies'
Add-Line -Text 'AI & Architecture Security: AI Security Architecture, LLM Security, Prompt Injection Defense, Model Governance, Threat Modeling, Secure SDLC, API Security' -Size 10 -SpaceAfter 1
Add-Line -Text 'Application Security: SAST, DAST, OWASP Top 10, Penetration Testing, AppSec Testing, Security Architecture Reviews' -Size 10 -SpaceAfter 1
Add-Line -Text 'Cloud Security: AWS, Azure, GCP, IAM, Cloud Compliance, Security Posture Management, Encryption' -Size 10 -SpaceAfter 1
Add-Line -Text 'GRC & Compliance: ISO 27001, ISO 42001, SOC 2 Type II, PCI DSS, RBI Framework, ITGC' -Size 10 -SpaceAfter 1
Add-Line -Text 'Programming & Tools: Python, Java, SQL, Bash/Linux Scripting' -Size 10 -SpaceAfter 3

# ---------------- EXPERIENCE ----------------
Add-Heading 'Professional Experience'

Add-Role -Left 'Consultant - AI & Solution Security Architect' -Right 'March 2026 - Present'
Add-SubLine 'NuSummit Cybersecurity, Mumbai, India'
Add-Bullet 'Conducted network architecture and firewall security reviews (segmentation, VPNs, IDS/IPS), identifying critical security gaps and reducing misconfigurations by 30% across client environments.'
Add-Bullet 'Performed AI security architecture reviews for AI-integrated products, identifying prompt injection, model governance gaps, and LLM data privacy risks, reducing AI risk exposure by 35% across 5+ systems.'
Add-Bullet 'Assessed cloud security architecture across AWS, Azure, and GCP, uncovering IAM misconfigurations and encryption gaps and achieving 100% remediation compliance for 8+ critical applications.'
Add-Bullet 'Performed threat modeling for cloud infrastructure and LLM deployments, delivering prioritized remediation roadmaps that reduced risk exposure by 40% across critical workloads.'

Add-Role -Left 'Senior Security Analyst - Application Security Architect' -Right 'August 2023 - March 2026'
Add-SubLine 'Deloitte Touche Tohmatsu India LLP, Mumbai, India'
Add-Bullet 'Designed and reviewed application security architectures for 15+ enterprise clients across BFSI, Fintech, and Healthcare, reducing architectural risk exposure by 40%.'
Add-Bullet 'Executed application security testing (SAST, DAST, API security) across 20+ web and mobile applications, uncovering OWASP Top 10 vulnerabilities and reducing critical findings by 45%.'
Add-Bullet 'Performed AI security architecture reviews for AI/ML solutions, identifying model poisoning and adversarial input risks, delivering remediation roadmaps that reduced AI risk exposure by 35%.'
Add-Bullet 'Executed compliance audits aligned with ISO 27001, PCI DSS, SOC 2 Type II, and RBI guidelines, achieving 100% audit clearance for all BFSI clients.'

Add-Role -Left 'Security Analyst' -Right 'December 2022 - May 2023'
Add-SubLine 'SDI, Bhubaneswar, India'
Add-Bullet 'Performed application and network security testing across web, mobile, and client-server applications, uncovering 50+ vulnerabilities and reducing network risk exposure by 30%.'
Add-Bullet 'Collaborated with development teams to validate security fixes, achieving zero re-occurrence of identified vulnerabilities across all retested systems.'

# ---------------- PROJECTS ----------------
Add-Heading 'Key Projects'
Add-Line -Text 'Security Audit - Payment Gateway System' -Size 11 -Bold $true -SpaceBefore 4 -SpaceAfter 0
Add-Bullet 'Conducted penetration testing across 100+ endpoints, identifying and mitigating 15+ high-severity vulnerabilities and reducing overall attack surface by 35% with zero re-occurrence during retesting.'
Add-Line -Text 'ATM Security Audit - Logical & Physical Controls' -Size 11 -Bold $true -SpaceBefore 4 -SpaceAfter 0
Add-Bullet 'Assessed security controls for 50+ ATMs, achieving 100% RBI compliance and reducing fraud risk exposure by 20% through targeted mitigation recommendations.'

# ---------------- CERTIFICATIONS ----------------
Add-Heading 'Certifications'
Add-Line -Text 'ISO 42001:2023 AI Management System Lead Auditor  |  ISO 27001:2022 Lead Auditor  |  CEH Practical  |  Certified AppSec Practitioner (CAP)' -Size 10 -SpaceAfter 2

# ---------------- EDUCATION ----------------
Add-Heading 'Education'
Add-Role -Left 'B.Tech in Electronics & Communication Engineering' -Right '2019 - 2023'
Add-SubLine 'Silicon Institute of Technology, Bhubaneswar, India  |  CGPA: 8.2 / 10'

# Remove trailing empty paragraph so it doesn't push onto a second page
$rng = $doc.Paragraphs($doc.Paragraphs.Count).Range
$rng.Delete() | Out-Null

# Save as .docx
$wdFormatDocumentDefault = 16
$doc.SaveAs([ref]$outPath, [ref]$wdFormatDocumentDefault)

# Export PDF (ExportAsFixedFormat takes a plain string path)
$pdfPath = $outPath -replace '\.docx$', '.pdf'
$doc.ExportAsFixedFormat($pdfPath, 17)  # 17 = wdExportFormatPDF

$pages = $doc.ComputeStatistics(2)  # 2 = wdStatisticPages

$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Output "Saved: $outPath"
Write-Output "Saved: $pdfPath"
Write-Output "Pages: $pages"
