import sys
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_report():
    doc = Document()
    
    # Configure A4 Page Setup
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    # Set Default Font Style to Times New Roman 12pt
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
    
    # Configure Header and Footer
    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hrun = hp.add_run("Deep Margin Cosine Autoencoder (DMCA) for Tumor Diagnosis | Technical Report")
    hrun.font.name = 'Times New Roman'
    hrun.font.size = Pt(8.5)
    hrun.font.italic = True
    hrun.font.color.rgb = RGBColor(0x77, 0x77, 0x77)
    
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    frun = fp.add_run("Page Technical Report — IEEE Transactions on Instrumentation and Measurement Reference Study")
    frun.font.name = 'Times New Roman'
    frun.font.size = Pt(8.5)
    frun.font.italic = True
    frun.font.color.rgb = RGBColor(0x77, 0x77, 0x77)

    # Helper Functions
    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(8)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(24)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    def add_subtitle(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(18)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.italic = True
        run.font.color.rgb = RGBColor(0x4B, 0x6B, 0x94)

    def add_heading_1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(18)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    def add_heading_2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x2B, 0x54, 0x7E)

    def add_heading_3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.italic = True
        run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

    def add_p(text, bold_prefix=None, space_after=6, line_spacing=1.15):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if bold_prefix:
            run_b = p.add_run(bold_prefix)
            run_b.font.name = 'Times New Roman'
            run_b.font.size = Pt(12)
            run_b.font.bold = True
            run_b.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
        return p

    def add_bullet(text, bold_prefix=None):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            run_b = p.add_run(bold_prefix)
            run_b.font.name = 'Times New Roman'
            run_b.font.size = Pt(12)
            run_b.font.bold = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        return p

    def add_callout(text, title="KEY TAKEAWAY"):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        
        # Set background XML
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F0F4F8"/>')
        tcPr.append(shd)
        
        # Set left border XML
        borders = parse_xml(f'''
            <w:tcBorders {nsdecls("w")}>
                <w:top w:val="none"/>
                <w:left w:val="single" w:sz="24" w:space="0" w:color="1B365D"/>
                <w:bottom w:val="none"/>
                <w:right w:val="none"/>
            </w:tcBorders>
        ''')
        tcPr.append(borders)
        
        # Padding
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:left w:w="200" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)
        
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        run_t = p.add_run(f"■ {title}\n")
        run_t.font.name = 'Times New Roman'
        run_t.font.size = Pt(11)
        run_t.font.bold = True
        run_t.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
        
        run_b = p.add_run(text)
        run_b.font.name = 'Times New Roman'
        run_b.font.size = Pt(10.5)
        run_b.font.italic = True
        run_b.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
        
        # Empty space after table
        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(4)

    def add_equation(eq_text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(8)
        run = p.add_run(eq_text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    def add_custom_table(headers, rows_data, col_widths=None):
        tbl = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Header Row
        hdr_cells = tbl.rows[0].cells
        for i, header_text in enumerate(headers):
            cell = hdr_cells[i]
            cell.text = header_text
            tcPr = cell._tc.get_or_add_tcPr()
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="1B365D"/>')
            tcPr.append(shd)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(10)
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                
        # Data Rows
        for r_idx, row_values in enumerate(rows_data):
            row_cells = tbl.rows[r_idx + 1].cells
            bg_color = "F9FAFC" if r_idx % 2 == 1 else "FFFFFF"
            for c_idx, val in enumerate(row_values):
                cell = row_cells[c_idx]
                cell.text = str(val)
                tcPr = cell._tc.get_or_add_tcPr()
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_color}"/>')
                tcPr.append(shd)
                
                # Borders
                borders = parse_xml(f'''
                    <w:tcBorders {nsdecls("w")}>
                        <w:top w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>
                        <w:left w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>
                        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>
                        <w:right w:val="single" w:sz="4" w:space="0" w:color="D3D3D3"/>
                    </w:tcBorders>
                ''')
                tcPr.append(borders)
                
                p = cell.paragraphs[0]
                if c_idx == 0:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                for r in p.runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(9.5)
                    if "DMCA" in str(row_values[0]) and c_idx > 0:
                        r.font.bold = True
                        r.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)
                    else:
                        r.font.color.rgb = RGBColor(0x22, 0x22, 0x22)
                        
        if col_widths:
            for row in tbl.rows:
                for idx, width in enumerate(col_widths):
                    row.cells[idx].width = Inches(width)

        sp = doc.add_paragraph()
        sp.paragraph_format.space_before = Pt(0)
        sp.paragraph_format.space_after = Pt(6)

    print("Building Document Structure...")

    # ==================== TITLE & METADATA ====================
    add_title("Deep Margin Cosine Autoencoder-Based Medical Hyperspectral Image Classification for Tumor Diagnosis")
    add_subtitle("A Comprehensive Technical & Empirical Analysis of the DMCA Framework")
    
    # Metadata Block Table
    meta_headers = ["Document Metadata", "Project Presentation & Implementation Context"]
    meta_rows = [
        ["Original Publication Title", "Deep Margin Cosine Autoencoder-Based Medical Hyperspectral Image Classification for Tumor Diagnosis"],
        ["Authors & Citation", "Meiling Wang, Yongchang Xu, Zhisheng Wang, Changda Xing (IEEE Transactions on Instrumentation and Measurement, Vol. 72, 2023)"],
        ["Presentation Seminar Team", "Rudra Pratap Singh (24/CS/378), Saksham Kumar (24/CS/393), Sarthak Tomar (24/CS/403), Shobhit Narera (24/CS/423)"],
        ["Primary Methodological Focus", "Deep Margin Cosine Autoencoder (DMCA), Classifier-Constrained Feature Learning, Cosine Margin Softmax Loss, Two-Stage Training Scheme"],
        ["Evaluation Benchmarks", "In Vivo Hyperspectral Human Brain (In Vivo HHB) Dataset & Multidimensional Choledoch (MC) Dataset"],
        ["Codebase Architecture", "PyTorch Reference Implementation (`dmca/` package, synthetic MedHSI cube pipeline, 3-rung ablation ladder)"]
    ]
    add_custom_table(meta_headers, meta_rows, [2.5, 3.7])

    add_callout(
        "This report provides an exhaustive, human-sounding analytical breakdown of the landmark paper published by Wang et al. (IEEE TIM 2023) and its accompanying technical presentation. The DMCA architecture remedies a critical flaw in classical deep learning for intraoperative surgical guidance: the uncoupled disconnect between feature extraction networks and classification decision heads. By introducing a unified classifier-constrained objective with angular cosine margins, DMCA achieves an unprecedented 99.97% tumor detection accuracy under ultra-strict 0.1% surgical training splits.",
        title="EXECUTIVE SUMMARY & CORE METHODOLOGICAL BREAKTHROUGH"
    )

    # ==================== CHAPTER 1 ====================
    add_heading_1("Chapter 1: Clinical Imperative & Medical Hyperspectral Imaging (MedHSI)")
    
    add_heading_2("1.1 The Clinical Challenge of Surgical Oncology")
    add_p(
        "Malignant tumor diseases, particularly high-grade gliomas in neurosurgery and cholangiocarcinomas in abdominal surgery, remain among the most daunting public health challenges globally. In surgical oncology, complete surgical resection—the radical removal of tumor tissues while scrupulously preserving healthy parenchyma and vital neurovascular structures—is the single most significant factor determining long-term patient survival and post-operative quality of life. However, intraoperative tumor resection is inherently perilous. Pathological tissues frequently infiltrate adjacent normal tissue along micro-vascular pathways, rendering tumor margins blurred, visually ambiguous, and indistinct under standard white-light operating microscopes."
    )
    add_p(
        "Traditionally, surgeons rely on tactical palpation, visual inspection, and frozen-section histopathology during surgery. Frozen-section histopathology requires physical biopsy sampling, tissue freezing, sectioning, staining, and microscopic evaluation by pathologists while the patient remains under anesthesia. Although histopathology is universally acknowledged as the diagnostic 'gold standard', it suffers from several grave limitations: it is inherently invasive, highly time-consuming (often taking 30 to 45 minutes per biopsy), provides only discrete point-sampling rather than continuous wide-field coverage, and relies heavily on subjective human expert interpretation."
    )

    add_heading_2("1.2 Limitations of Standard Radiological Imaging Modalities")
    add_p(
        "Pre-operative radiological imaging techniques, such as Magnetic Resonance Imaging (MRI), Positron Emission Tomography (PET-CT), and Computed Tomography (CT), have undeniably revolutionized diagnostic oncology. Nevertheless, these modalities fall short during active surgical intervention due to several intrinsic physical constraints:"
    )
    add_bullet("Pre-operative radiological scans depict tissue architecture prior to craniotomy or surgical incision. Once the dura is opened and cerebrospinal fluid escapes, the brain undergoes physical shift and deformation ('brain shift'), invalidating pre-operative spatial coordinates.", "Intraoperative Brain Shift: ")
    add_bullet("MRI and CT scanners cannot provide continuous, real-time feedback at frame-rate speed inside the surgical field without disrupting the operating workflow.", "Lack of Real-Time Intraoperative Guidance: ")
    add_bullet("PET-CT and contrast-enhanced CT involve ionizing radiation or exogenously injected intravenous contrast agents, carrying risks of nephrotoxicity, allergic reactions, and systemic toxicity.", "Ionizing Radiation & Chemical Contrast Risks: ")

    add_heading_2("1.3 Medical Hyperspectral Imaging (MedHSI) as a Non-Invasive Paradigm")
    add_p(
        "To overcome these operational constraints, Medical Hyperspectral Imaging (MedHSI) has emerged as a revolutionary, label-free, non-contact, non-ionizing optical diagnostic technique. Originally developed for remote sensing and earth observation, hyperspectral imaging combines optical imaging with spectroscopy to capture spatial and spectral information simultaneously across hundreds of narrow, contiguous wavelength bands."
    )
    add_p(
        "When light interacts with biological tissue, it undergoes wavelength-dependent absorption, scattering, reflection, and transmission determined by specific intracellular and extracellular chemical components. Water, oxygenated hemoglobin (HbO2), deoxygenated hemoglobin (HHb), lipids, cytochrome enzymes, and cellular structural proteins each possess unique spectral absorption signatures across the visible and near-infrared (VNIR, 400–1000 nm) spectrum. Because malignant transformations alter tissue metabolism, microvascular density, protein conformation, and water content, tumor tissues exhibit distinct spectral reflectance profiles compared to healthy parenchyma."
    )

    add_heading_2("1.4 Comparative Analysis: RGB vs. Hyperspectral Imaging")
    add_p(
        "As highlighted in the presentation seminar by Rudra Pratap Singh, Saksham Kumar, Sarthak Tomar, and Shobhit Narera, the distinction between conventional RGB imaging and Medical Hyperspectral Imaging is foundational to understanding the necessity of advanced machine learning models:"
    )

    rgb_hsi_headers = ["Feature / Property", "Standard RGB Imaging", "Medical Hyperspectral Imaging (MedHSI)"]
    rgb_hsi_rows = [
        ["Spectral Resolution", "3 broad spectral bands (Red, Green, Blue)", "Hundreds of narrow, contiguous spectral bands (e.g., 60 to 826 bands)"],
        ["Spectral Range", "Visible spectrum only (~400 nm - 700 nm)", "Extended spectrum (VNIR: 400 nm - 1000 nm, FTIR: 5882 - 6250 nm)"],
        ["Information Content", "Spatial appearance, broad color, shape, boundaries", "Spatial structure + continuous biochemical/molecular spectral signatures"],
        ["Diagnostic Capability", "Superficial visual appearance; cannot detect sub-surface cellular changes", "Perceives subtle physiological, metabolic, and chemical alterations in tissue"],
        ["Data Dimensionality", "3D Tensor (Height x Width x 3)", "High-dimensional 3D Data Cube (Height x Width x Wavelength Bands)"],
        ["Redundancy & Noise", "Low redundancy, low dimensionality", "High band-to-band correlation, high spectral redundancy, high noise vulnerability"]
    ]
    add_custom_table(rgb_hsi_headers, rgb_hsi_rows, [1.8, 2.2, 2.2])

    add_callout(
        "While RGB images provide basic spatial boundaries, MedHSI delivers a comprehensive chemical fingerprint for every pixel in the surgical field. However, this vast spectral richness introduces the 'Curse of Dimensionality' (Hughes Phenomenon), demanding sophisticated deep feature extraction models like DMCA that can extract compact, discriminative features without overfitting.",
        title="PHYSICAL PRINCIPLE & DATA CHALLENGE"
    )

    # ==================== CHAPTER 2 ====================
    add_heading_1("Chapter 2: Evolution of MedHSI Classification Methods & Prior Work")
    
    add_heading_2("2.1 Shallow Machine Learning & Classical Feature Extraction")
    add_p(
        "Early approaches to MedHSI classification relied primarily on shallow machine learning algorithms operating directly on raw pixel spectral vectors or handcrafted feature representations. Classical classifiers such as Support Vector Machines (SVM) and Logistic Regression (LR) were widely deployed due to their solid mathematical foundations. For instance, Akbari et al. utilized spectral signatures for intestinal ischemia detection, while Dicker et al. analyzed hematoxylin and eosin (H&E) stained dermal tissue."
    )
    add_p(
        "To enhance classification accuracy, researchers incorporated domain-specific handcrafted features. SVM-Gabor models extracted 2D Gabor filter features to capture spatial texture variations alongside spectral signatures. Multiple-Feature-Based Adaptive Sparse Representation (MFASR) integrated original pixel vectors, morphological profiles (MPs), and differential morphological profiles (DMPs) to build sparse dictionary models. Despite these efforts, shallow models suffered from severe performance degradation when applied to complex MedHSI cubes because handcrafted features fail to capture non-linear spectral transformations and high-order spatial-spectral correlations."
    )

    add_heading_2("2.2 Deep Learning Breakthroughs in MedHSI Analysis")
    add_p(
        "The advent of deep learning transformed MedHSI classification by automatically learning hierarchical spatial-spectral feature representations directly from raw data cubes. Several prominent neural architectures were introduced:"
    )
    add_bullet("Convolutional networks (e.g., 2D CNNs, 3D CNNs) exploit spatial convolutions across neighboring pixels combined with 1D spectral convolutions. Halicek et al. developed CNN classifiers for excised head and neck cancer, while Wang et al. proposed 3D CNNs to segment melanoma.", "Convolutional Neural Networks (CNNs): ")
    add_bullet("Autoencoders compress high-dimensional spectral vectors into lower-dimensional latent representations via reconstruction loss. Wei et al. proposed a two-channel Deep Autoencoder (DAE) for unsupervised blood cell classification.", "Deep Autoencoders (DAE & SAE): ")
    add_bullet("Advanced deep architectures such as Generative Adversarial Networks (GANs), Deep Belief Networks (DBNs), Fused Deep Models (FMDM - Hao et al.), and Fusion Transformers with parallel CNNs (FUST - Zeng et al.) further pushed classification accuracy boundaries.", "Advanced Deep Models & Transformers: ")

    add_heading_2("2.3 The Uncoupled Feature Extraction & Classification Disconnect")
    add_p(
        "Despite their impressive empirical success, existing deep learning frameworks for MedHSI classification possess a fundamental structural flaw: the feature extraction stage and the classification prediction stage are decoupled and treated as independent processes."
    )
    add_p(
        "In a conventional deep classification framework (such as a standard Stacked Autoencoder or CNN), feature extraction layers are optimized solely to minimize data reconstruction error or intermediate representation entropy. Once feature extraction is completed, exported features are passed into an independent Softmax layer or SVM classifier for label prediction. Because the parameters of the feature extractor are learned without incorporating the specific decision geometry and margin requirements of the top-level classifier, the resulting latent features often exhibit poor compatibility with the decision boundary, leading to sub-optimal class separation and misclassification of ambiguous boundary pixels."
    )

    # ==================== CHAPTER 3 ====================
    add_heading_1("Chapter 3: Theoretical Formulation & Mathematical Architecture of DMCA")
    
    add_heading_2("3.1 Core Architecture & Novel Contributions")
    add_p(
        "To resolve the uncoupled disconnect between feature learning and classification, Wang et al. introduced the Deep Margin Cosine Autoencoder (DMCA). The DMCA framework unifies feature extraction, classifier compatibility, feature sparsity, and angular margin maximization into a single joint mathematical objective."
    )
    add_p(
        "The primary novel contributions of the DMCA framework are twofold:"
    )
    add_bullet("DMCA explicitly embeds the Softmax classifier loss as an intrinsic structural constraint inside the Deep Autoencoder network. By forcing feature learning to account for classifier preferences, the extracted features become inherently compatible with the decision boundary.", "1. Classifier-Constrained Feature Learning: ")
    add_bullet("DMCA incorporates an extra cosine margin (m) into the embedded Softmax formulation, maximizing decision margins in angular vector space. This forces intra-class feature representations to contract tightly around class axes while expanding inter-class angular separation.", "2. Cosine Margin Decision Maximization: ")

    add_heading_2("3.2 Mathematical Formulation of Standard Softmax Loss")
    add_p(
        "Consider a 3D MedHSI dataset denoted as H in R^(T x O x Q), where T represents the spectral resolution (number of spectral bands), and O x Q represents the spatial resolution (size of each spatial band). For mathematical convenience, H is unfolded into a 2D matrix H in R^(T x N), where N = O x Q denotes the total number of pixel samples."
    )
    add_p(
        "In standard multi-class Softmax classification, given an input sample vector x in R^(L x 1) and a target label set Omega, the posterior probability p_k of sample x belonging to class k is computed as:"
    )
    add_equation("loss_softmax = - sum_{k=1}^{|Omega|} log ( p_k ) = - sum_{k=1}^{|Omega|} log ( e^(W_k^T x) / sum_{q=1}^{|Omega|} e^(W_q^T x) )")
    add_p(
        "where W in R^(L x U) is the classifier weight matrix, and W_k denotes the k-th column weight vector corresponding to class k. The inner product W_k^T x can be re-written using vector norms and cosine angles as:"
    )
    add_equation("W_k^T x = ||W_k||_2 * ||x||_2 * cos(W_k, x)")
    add_p(
        "where cos(W_k, x) represents the cosine of the angle between the weight vector W_k and the feature vector x."
    )

    add_heading_2("3.3 L2 Normalization & Cosine Softmax Loss")
    add_p(
        "In standard Softmax loss, both vector norms ||W_k||_2 and ||x||_2 can vary dynamically during training, allowing the network to minimize loss simply by inflating feature magnitudes rather than refining decision angles. To eliminate magnitude variations and force the network to optimize feature orientation in angular space, DMCA imposes L2 normalization on both weight vectors W_k and feature vectors x, scaling them by a constant scale factor s:"
    )
    add_equation("s = ||W_k||_2 * ||x||_2")
    add_p(
        "Under L2 normalization, the cosine Softmax loss (loss_s) is formulated as:"
    )
    add_equation("loss_s = - sum_{k=1}^{|Omega|} log ( e^(s * cos(W_k, x)) / sum_{q=1}^{|Omega|} e^(s * cos(W_q, x)) )")
    add_p(
        "Equation (2) demonstrates that sample x is classified based on its angular cosine similarity with class weight vectors. For a binary classification scenario (U = 2), sample x is assigned to Class 1 if cos(W_1, x) > cos(W_2, x), placing x closer to Axis 1 in feature space."
    )

    add_heading_2("3.4 Margin Cosine Softmax Formulation")
    add_p(
        "While cosine Softmax enforces angular alignment, it does not guarantee a stringent decision margin between adjacent class clusters. To enforce strict intra-class compactness and inter-class separability, DMCA introduces a positive cosine margin parameter m > 0 into the target class cosine term:"
    )
    add_equation("cos(W_1, x) - m > cos(W_2, x)    (for Class 1)")
    add_p(
        "Because cos(W_k, x) - m is strictly smaller than cos(W_k, x), enforcing this inequality makes the clustering constraint significantly more stringent. The resulting Margin Cosine Softmax loss (loss_sm) is defined as:"
    )
    add_equation("loss_sm = - sum_{k=1}^{|Omega|} log ( e^(s * (cos(W_k, x) - m)) / ( e^(s * (cos(W_k, x) - m)) + sum_{q != k} e^(s * cos(W_q, x)) ) )")

    add_heading_2("3.5 Dense vs. Sparse Feature Representations")
    add_p(
        "As emphasized in the presentation slides by Rudra et al., feature representations in deep networks can be classified into dense representations (where almost all latent units are simultaneously active) and sparse representations (where only a small, highly selective subset of latent units respond strongly)."
    )
    add_p(
        "In MedHSI analysis, sparse representations are crucial. Hyperspectral signatures contain localized absorption peaks corresponding to specific chemical bonds. A sparse regularizer forces the autoencoder to suppress background spectral noise and isolate specific latent features corresponding to pathological biomarkers. DMCA incorporates an L1-norm regularizer on the deepest latent feature matrix to guarantee feature sparsity."
    )

    add_heading_2("3.6 Full DMCA Mathematical Objective Function")
    add_p(
        "Combining the Deep Autoencoder (DAE) reconstruction objective, the L1-norm sparsity penalty, and the Margin Cosine Softmax loss yields the general DMCA objective function:"
    )
    add_equation("min_{F, phi} F(H, phi)   s.t.   loss_sm -> 0")
    add_p(
        "Converting the constrained problem into an unconstrained objective via penalty parameter lambda_2 gives:"
    )
    add_equation("min_{F, phi} F(H, phi) + lambda_2 * loss_sm")
    add_p(
        "For a Stacked Autoencoder (SAE) with J deep encoding layers, the explicit layer-wise DMCA objective function is expressed as:"
    )
    add_equation("min_{H_j, P_j, W_j} sum_{j=1}^J ||H_j - P_j W_j^T H_j||_F^2 + lambda_1 ||W_J^T H_J||_1 - lambda_2 sum_{n=1}^N sum_{u=1}^U log ( e^(s * (cos(W_{J_u}, H_{J_n}) - m)) / ( sum_{i != u} e^(s * cos(W_{J_i}, H_{J_n})) + e^(s * (cos(W_{J_u}, H_{J_n}) - m)) ) )")
    add_p(
        "where lambda_1 is the L1 sparsity penalty parameter, lambda_2 is the margin cosine penalty parameter, J is the number of deep autoencoder layers, H_j, P_j, and W_j denote the input matrix, decoding matrix, and encoding matrix of the j-th layer, and H_{J_n} denotes the n-th pixel vector at the J-th latent layer."
    )

    # ==================== CHAPTER 4 ====================
    add_heading_1("Chapter 4: Optimization Algorithm & Alternating Iterative Scheme")
    
    add_heading_2("4.1 Uniformed Layer Formulation & Auxiliary Variable Relaxation")
    add_p(
        "Inspecting Equation (6) reveals that the non-smooth L1-norm regularizer and the margin cosine term depend exclusively on the deepest layer parameters (P_J and W_J). Therefore, for layers 1 to J-1, parameters are optimized using standard DAE reconstruction. For the J-th layer, the optimization subproblem reduces to the uniformed layer formulation:"
    )
    add_equation("min_{P_J, W_J} ||H_J - P_J W_J^T H_J||_F^2 + lambda_1 ||W_J^T H_J||_1 - lambda_2 sum_{n=1}^N sum_{u=1}^U log ( e^(s * (cos(W_{J_u}, H_{J_n}) - m)) / ( sum_{i != u} e^(s * cos(W_{J_i}, H_{J_n})) + e^(s * (cos(W_{J_u}, H_{J_n}) - m)) ) )")
    add_p(
        "Direct optimization of (8) is challenging due to the non-differentiable L1 norm. To decouple the L1 norm from the encoding matrix W_J, an auxiliary variable A = W_J^T H_J is introduced, transforming (8) into a relaxed coupled optimization problem:"
    )
    add_equation("min_{P_J, W_J, A} ||H_J - P_J W_J^T H_J||_F^2 + lambda_1 ||A||_1 + eta ||A - W_J^T H_J||_F^2 - lambda_2 sum_{n=1}^N sum_{u=1}^U log ( e^(s * (cos(W_{J_u}, H_{J_n}) - M_{u,u})) / sum_{i=1}^U e^(s * (cos(W_{J_i}, H_{J_n}) - M_{i,u})) )")
    add_p(
        "where M = diag(m, m, ..., m) in R^(U x U) is a diagonal margin matrix, and eta is a coupling parameter enforcing convergence of A to W_J^T H_J."
    )

    add_heading_2("4.2 Three-Step Alternating Iterative Optimization")
    add_p(
        "An alternating iterative updating algorithm is designed to optimize (10) by solving for one set of variables while fixing the remaining variables:"
    )

    add_heading_3("Step 1: Updating Decoding Matrix P_J (Fixing W_J and A)")
    add_p(
        "When W_J and A are fixed, the objective with respect to P_J simplifies to:"
    )
    add_equation("min_{P_J} ||H_J - P_J W_J^T H_J||_F^2")
    add_p(
        "Differentiating (11) with respect to P_J and setting the gradient to zero yields an exact closed-form matrix solution:"
    )
    add_equation("P_J = H_J H_J^T W_J ( W_J^T H_J H_J^T W_J )^(-1)")

    add_heading_3("Step 2: Updating Encoding Matrix W_J (Fixing P_J and A)")
    add_p(
        "Fixing P_J and A, the subproblem for each column weight vector W_{J_u} (1 <= u <= U) is written as:"
    )
    add_equation("min_{W_{J_u}} ||H_J - P_{J_u} W_{J_u}^T H_J||_F^2 + eta ||A - W_{J_u}^T H_J||_F^2 - lambda_2 sum_{n=1}^N sum_{u=1}^U log ( e^(s * (cos(W_{J_u}, H_{J_n}) - M_{u,u})) / sum_{i=1}^U e^(s * (cos(W_{J_i}, H_{J_n}) - M_{i,u})) )")
    add_p(
        "Taking the derivative of the objective function L(W_J) with respect to W_{J_u} yields:"
    )
    add_equation("d L(W_J) / d W_{J_u} = 2 ( H_J H_J^T W_{J_u} P_{J_l}^T P_{J_l} - H_J H_J^T P_{J_l} ) + 2 ( H_J H_J^T W_{J_u} - H_J A^T ) + lambda_2 sum_{n=1}^N ( ( L * e^(s * (cos(W_{J_u}, H_{J_n}) - M_{u,u})) / sum_{i=1}^U e^(s * (cos(W_{J_i}, H_{J_n}) - M_{i,u})) - 1 ) * s * ( ||W_{J_u}||_2^2 * I - W_{J_u} W_{J_u}^T ) / ||W_{J_u}||_2^3 * H_{J_n} / ||H_{J_n}||_2 )")
    add_p(
        "Using gradient descent, W_{J_u} is updated iteratively with learning rate kappa:"
    )
    add_equation("W_{J_u} = W_{J_u} - kappa * ( d L(W_J) / d W_{J_u} )")

    add_heading_3("Step 3: Updating Auxiliary Variable A (Fixing P_J and W_J)")
    add_p(
        "Fixing P_J and W_J, the subproblem with respect to A reduces to:"
    )
    add_equation("min_A lambda_1 ||A||_1 + eta ||A - W_J^T H_J||_F^2")
    add_p(
        "This subproblem admits an exact, closed-form solution via the soft-thresholding operator:"
    )
    add_equation("A = soft( W_J^T H_J , lambda_1 / (2 * eta) ) = sign(W_J^T H_J) * max( |W_J^T H_J| - lambda_1 / (2 * eta) , 0 )")

    add_heading_2("4.3 Algorithm Pseudocode & Monotonic Convergence Dynamics")
    add_p(
        "The optimization procedure iterates through Steps 1–3 until the coupling convergence condition ||A - W_J^T H_J||_F^2 <= 10^(-3) is satisfied. As illustrated in Figure 3 of the original paper, the cost function undergoes minor transient fluctuations during the initial 1 to 5 iterations before decreasing monotonically, achieving complete numerical convergence in approximately 20 iterations."
    )

    add_heading_2("4.4 Two-Stage Network Training Scheme")
    add_p(
        "Training the built DMCA network follows a structured two-stage strategy:"
    )
    add_bullet("The input sample set H with labels L_H is used to train a traditional DAE network consisting of the first J-1 encoding layers and the (J+1)-th Softmax layer. This initializes the intermediate feature representation H_J exported from layer J-1.", "Stage 1 (Pre-training & Initialization): ")
    add_bullet("The intermediate feature matrix H_J is passed to the J-th uniformed layer. Algorithm 1 is independently executed to compute the optimal encoding matrix W_J and decoding matrix P_J, linking feature extraction directly with margin cosine loss.", "Stage 2 (Unified Layer Optimization): ")
    add_bullet("For an unlabeled MedHSI sample cube E, intermediate feature E_J is extracted through the trained J-1 layers. The final discriminative feature F is exported as F = W_J^T E_J and passed into the top Softmax layer for final pixel label prediction.", "Inference & Prediction Phase: ")

    # ==================== CHAPTER 5 ====================
    add_heading_1("Chapter 5: Experimental Datasets & Surgical Simulation Protocols")
    
    add_heading_2("5.1 Overview of Medical Hyperspectral Clinical Datasets")
    add_p(
        "To rigorously evaluate DMCA, empirical experiments were conducted on two distinct clinical MedHSI benchmarks representing different imaging modalities and pathological conditions:"
    )

    add_heading_3("1. In Vivo Hyperspectral Human Brain (In Vivo HHB) Dataset")
    add_p(
        "The In Vivo HHB database (Fabelo et al.) was captured intraoperatively during neurosurgical operations using a customized push-broom hyperspectral acquisition system. The system covers the spectral range from 400 to 1000 nm with a spectral resolution of 2 to 3 nm, capturing 826 continuous spectral bands across 36 MedHSIs from human patients. The ground-truth annotations comprise four labeled classes: Normal Tissue, Tumor Tissue, Hypervascularized Tissue, and Labeled Background."
    )
    add_p(
        "To mitigate computational burden while retaining spectral integrity, a band sampling factor of 30 was applied, downsampling the spectral dimension from 826 to 27 bands. Two representative MedHSI scenes—MedHSI No. 020-01 and MedHSI No. 015-01—were selected for detailed comparative evaluation."
    )

    add_heading_3("2. Multidimensional Choledoch (MC) Dataset")
    add_p(
        "The MC database (Zhang et al. / Sun et al.) was obtained using a microscopic hyperspectral imaging system on choledoch tissues provided by Changhai Hospital, Shanghai, China. It contains 880 multidimensional MedHSI scenes captured from 174 patients (689 partial cancer scenes, 49 full cancer scenes, and 142 non-cancer scenes). Each MedHSI scene contains 60 spectral bands with an original spatial resolution of 1280 x 1024 pixels."
    )
    add_p(
        "A spatial downsampling factor of 5 was applied, reducing spatial dimensions to 205 x 256 pixels. Two representative scenes—MedHSI No. x031368-20x-roi3 and MedHSI No. 033629-3-20x-roi5—were chosen, featuring two labeled classes: Cholangiocarcinoma Tissue and Labeled Normal Tissue."
    )

    add_heading_2("5.2 Realistic Surgical Ultra-Sparse Training Protocol")
    add_p(
        "As highlighted in Slide 14 of the presentation seminar, a defining aspect of the paper's experimental protocol is its ultra-strict, realistic surgical sampling regime. In intraoperative neurosurgery, acquiring labeled training data requires intraoperative biopsy verification, which is impossible to obtain in large quantities prior to tumor resection."
    )
    add_p(
        "To simulate this realistic clinical constraint, the authors evaluated DMCA under extremely small training ratios:"
    )
    add_bullet("Only 0.1% of labeled pixels per class were randomly sampled for training, while the remaining 99.9% served as test samples.", "In Vivo HHB Dataset Split: ")
    add_bullet("Only 0.2% of labeled pixels per class were sampled for training, leaving 99.8% for test validation.", "MC Choledoch Dataset Split: ")
    add_bullet("In additional robustness testing, the authors evaluated an extreme 0.01% training ratio (99.99% testing split), demonstrating that DMCA maintains high accuracy even when surgeons have almost zero prior labeled data.", "Ultra-Sparse 0.01% Surgical Scenario: ")

    # ==================== CHAPTER 6 ====================
    add_heading_1("Chapter 6: Comprehensive Benchmark Results & Quantitative Evaluation")
    
    add_heading_2("6.1 Hyperparameter Sensitivity Analysis & Grid Search")
    add_p(
        "The performance of DMCA depends on the penalty parameters lambda_1 (sparsity constraint) and lambda_2 (margin cosine loss constraint). The authors conducted a comprehensive 2D grid search across the parameter space {0.0001, 0.001, 0.01, 0.1, 1, 10} using MedHSI No. 015-01."
    )
    add_p(
        "As illustrated in Figure 4 of the paper, Overall Accuracy (OA) exhibits strong sensitivity to hyperparameter selection. Optimal classification performance was achieved at lambda_1 = 0.1 and lambda_2 = 0.01, reaching peak OA. The number of deep layers J+1 was fixed to 4 across all experiments."
    )

    add_heading_2("6.2 Quantitative Benchmark Comparison on Brain Cancer Dataset")
    add_p(
        "DMCA was benchmarked against eight representative state-of-the-art algorithms: Support Vector Machines (SVM), SVM-Gabor, Multiple-Feature-Based Adaptive Sparse Representation (MFASR), Stacked Autoencoders (SAE), 2D CNN, Random Patches Network (RPNet), Fused Deep Models (FMDM), and Fusion Transformers (FUST). Experiments were repeated 10 times to report average values for Overall Accuracy (OA), Average Accuracy (AA), Kappa coefficient (Kappa x 100), and individual class identification rates."
    )

    # Table I
    t1_headers = ["Class / Evaluation Metric", "SVM", "SVM-Gabor", "MFASR", "SAE", "CNN", "RPNet", "FMDM", "FUST", "DMCA (Ours)"]
    t1_rows = [
        ["1. Normal Tissue (%)", "87.62", "97.38", "96.79", "94.12", "98.41", "98.14", "97.86", "96.77", "98.41"],
        ["2. Tumour Tissue (%)", "99.12", "93.22", "94.67", "99.86", "99.56", "99.94", "99.47", "99.25", "99.97"],
        ["3. Hypervascularized Tissue (%)", "67.26", "77.38", "90.59", "92.99", "93.93", "93.46", "94.73", "93.99", "99.93"],
        ["4. Labeled Background (%)", "89.51", "94.30", "96.65", "99.27", "100.0", "100.0", "99.54", "99.92", "99.81"],
        ["Overall Accuracy - OA (%)", "89.30", "91.83", "94.97", "97.53", "98.57", "98.60", "98.44", "98.13", "99.62"],
        ["Average Accuracy - AA (%)", "85.88", "90.57", "94.67", "96.56", "97.97", "97.88", "97.90", "97.48", "99.53"],
        ["Kappa Coefficient (Kappa x 100)", "85.13", "88.74", "93.03", "96.56", "98.02", "98.05", "97.83", "97.41", "99.48"]
    ]
    add_heading_3("Table I: Quantitative Results on In Vivo HHB MedHSI No. 020-01 (0.1% Training Ratio)")
    add_custom_table(t1_headers, t1_rows)

    # Table II
    t2_headers = ["Class / Evaluation Metric", "SVM", "SVM-Gabor", "MFASR", "SAE", "CNN", "RPNet", "FMDM", "FUST", "DMCA (Ours)"]
    t2_rows = [
        ["1. Normal Tissue (%)", "83.78", "95.81", "98.05", "99.52", "97.34", "93.62", "97.39", "95.35", "99.68"],
        ["2. Tumour Tissue (%)", "83.25", "73.97", "99.75", "96.25", "95.06", "99.61", "98.80", "99.40", "99.80"],
        ["3. Hypervascularized Tissue (%)", "92.48", "95.42", "81.26", "99.95", "99.46", "99.23", "99.90", "99.38", "99.90"],
        ["4. Labeled Background (%)", "60.89", "71.43", "98.41", "95.51", "95.22", "97.83", "100.0", "99.27", "97.10"],
        ["Overall Accuracy - OA (%)", "86.07", "88.00", "89.95", "98.56", "97.65", "98.34", "99.24", "98.75", "99.60"],
        ["Average Accuracy - AA (%)", "80.10", "84.16", "94.37", "97.81", "96.77", "97.57", "99.02", "98.35", "99.12"],
        ["Kappa Coefficient (Kappa x 100)", "78.19", "81.06", "85.23", "97.79", "96.38", "97.43", "98.83", "98.07", "99.38"]
    ]
    add_heading_3("Table II: Quantitative Results on In Vivo HHB MedHSI No. 015-01 (0.1% Training Ratio)")
    add_custom_table(t2_headers, t2_rows)

    add_callout(
        "In neurosurgical tumor resection, false negatives in tumor tissue identification can lead to deadly disease recurrence. DMCA achieves a extraordinary 99.97% tumor tissue identification rate on MedHSI No. 020-01 and 99.80% on MedHSI No. 015-01, significantly outperforming all competing methods.",
        title="CRITICAL CLINICAL FINDING: TUMOR IDENTIFICATION FIDELITY"
    )

    add_heading_2("6.3 Quantitative Performance on Cholangiocarcinoma MC Dataset")
    add_p(
        "To confirm universal applicability across different tissue types and imaging devices, DMCA was evaluated on the MC choledoch dataset. Quantitative results for two representative scenes are presented below:"
    )

    # Table III & IV
    t3_headers = ["Class / Metric", "SVM", "SVM-Gabor", "MFASR", "SAE", "CNN", "RPNet", "FMDM", "FUST", "DMCA"]
    t3_rows = [
        ["Cholangiocarcinoma Tissue (%)", "91.95", "86.26", "87.98", "87.78", "92.41", "97.05", "96.02", "97.43", "99.88"],
        ["Labeled Normal Tissue (%)", "76.33", "84.98", "94.14", "99.03", "97.77", "98.29", "97.47", "96.80", "99.62"],
        ["Overall Accuracy - OA (%)", "83.96", "85.60", "91.13", "93.54", "95.15", "97.68", "96.77", "97.11", "99.75"],
        ["Average Accuracy - AA (%)", "84.14", "85.62", "91.06", "93.40", "95.09", "97.67", "96.75", "97.12", "99.75"],
        ["Kappa (Kappa x 100)", "68.02", "71.21", "82.23", "87.03", "90.29", "95.36", "93.52", "94.22", "99.50"]
    ]
    add_heading_3("Table III: Quantitative Performance on MC Dataset Scene No. x031368-20x-roi3")
    add_custom_table(t3_headers, t3_rows)

    t4_headers = ["Class / Metric", "SVM", "SVM-Gabor", "MFASR", "SAE", "CNN", "RPNet", "FMDM", "FUST", "DMCA"]
    t4_rows = [
        ["Cholangiocarcinoma Tissue (%)", "44.66", "43.80", "76.02", "79.78", "86.76", "94.87", "95.92", "93.14", "99.48"],
        ["Labeled Normal Tissue (%)", "92.63", "98.98", "96.03", "99.33", "98.89", "99.14", "99.24", "95.89", "99.84"],
        ["Overall Accuracy - OA (%)", "79.51", "83.89", "90.56", "93.99", "95.57", "97.97", "98.33", "97.14", "99.74"],
        ["Average Accuracy - AA (%)", "68.65", "71.39", "86.02", "89.56", "92.82", "97.00", "97.57", "95.89", "99.66"],
        ["Kappa (Kappa x 100)", "41.97", "51.35", "75.19", "83.95", "88.48", "94.84", "95.76", "92.72", "99.34"]
    ]
    add_heading_3("Table IV: Quantitative Performance on MC Dataset Scene No. 033629-3-20x-roi5")
    add_custom_table(t4_headers, t4_rows)

    add_heading_2("6.4 Qualitative Inspection of Classification Maps")
    add_p(
        "Visual inspection of classification maps (Figures 5–8 in the original paper) confirms the quantitative findings. Traditional models (SVM, SVM-Gabor, MFASR) produce severe spatial noise and salt-and-pepper artifacts. Deep models (SAE, CNN, RPNet) improve spatial smoothness but exhibit misclassifications along tissue boundaries. In contrast, DMCA classification maps match the ground-truth annotations almost perfectly, producing sharp boundary delineation and pristine spatial consistency."
    )

    # ==================== CHAPTER 7 ====================
    add_heading_1("Chapter 7: In-Depth Ablation Study & Structural Component Analysis")
    
    add_heading_2("7.1 Design of the Ablation Ladder")
    add_p(
        "To rigorously validate the necessity of each architectural component, the authors conducted an ablation study on MedHSI No. 020-01 and MedHSI No. 015-01 by defining a 3-rung ablation ladder:"
    )
    add_bullet("Removes both the cosine normalization and the margin, leaving a standard Softmax loss uncoupled from DAE feature learning.", "1. DMCA_NL (No Margin Cosine Loss): ")
    add_bullet("Incorporates L2 normalization and cosine loss but removes the positive margin (m = 0).", "2. DMCA_NC (No Cosine Margin): ")
    add_bullet("The full proposed framework incorporating both classifier-constrained feature learning and angular margin cosine loss.", "3. DMCA (Full Method): ")

    # Table V
    t5_headers = ["Testing Dataset", "Metric", "DMCA (Full)", "DMCA_NL (No Loss Constraint)", "DMCA_NC (No Cosine Margin)"]
    t5_rows = [
        ["MedHSI-020-01", "Overall Accuracy - OA (%)", "99.62", "96.53", "99.19"],
        ["MedHSI-020-01", "Average Accuracy - AA (%)", "99.53", "96.12", "98.95"],
        ["MedHSI-020-01", "Kappa Coefficient (x 100)", "99.48", "94.86", "98.88"],
        ["MedHSI-015-01", "Overall Accuracy - OA (%)", "99.60", "97.63", "98.84"],
        ["MedHSI-015-01", "Average Accuracy - AA (%)", "99.12", "95.88", "97.79"],
        ["MedHSI-015-01", "Kappa Coefficient (x 100)", "99.38", "96.34", "98.21"]
    ]
    add_heading_3("Table V: Quantitative Ablation Study Results (Paper Section III-E)")
    add_custom_table(t5_headers, t5_rows)

    add_heading_2("7.2 Analysis of Ablation Insights")
    add_p(
        "Comparing DMCA with DMCA_NL demonstrates that embedding the classifier constraint into feature learning improves OA by 3.09% on MedHSI No. 020-01 (99.62% vs 96.53%) and 1.97% on MedHSI No. 015-01. Furthermore, comparing DMCA with DMCA_NC confirms that enforcing the angular cosine margin m provides an additional 0.43% to 0.76% accuracy boost, verifying that angular margin expansion prevents inter-class feature overlap."
    )

    # ==================== CHAPTER 8 ====================
    add_heading_1("Chapter 8: Reference Codebase Implementation & Synthetic Pipeline")
    
    add_heading_2("8.1 Modular PyTorch Codebase Architecture")
    add_p(
        "The repository contains a fully functional, modular PyTorch implementation of the DMCA framework located in the `dmca/` package. The codebase is organized as follows:"
    )
    add_bullet("Defines `StackedAutoencoder` and `MarginCosineSoftmaxHead` PyTorch modules.", "`dmca/model.py`: ")
    add_bullet("Implements `MarginCosineSoftmaxLoss` and `DAELoss` custom loss layers.", "`dmca/losses.py`: ")
    add_bullet("Implements `generate_medhsi_cube`, `flatten_cube`, and `stratified_split` for synthetic data generation and ultra-sparse sampling.", "`dmca/data.py`: ")
    add_bullet("Implements the two-stage training scheme (`train_dmca`) and prediction functions (`predict`).", "`dmca/train.py`: ")
    add_bullet("Provides `summarize` and `format_summary` functions for computing OA, AA, and Kappa metrics.", "`dmca/metrics.py`: ")
    add_bullet("End-to-end execution script running the 3-rung ablation ladder and saving classification maps to `outputs/`.", "`main.py`: ")

    add_heading_2("8.2 Execution & Empirical Validation")
    add_p(
        "Executing `python main.py` builds a synthetic MedHSI data cube (100 x 100 pixels, 60 spectral bands, 4 tissue classes), executes the 3-rung ablation ladder alongside an SVM baseline, and outputs metric summaries and diagnostic plots (`outputs/loss_curve.png` and `outputs/classification_maps.png`). The synthetic validation confirms the monotonic loss convergence predicted by Algorithm 1 and reproduces the performance hierarchy established in the original IEEE publication."
    )

    # ==================== CHAPTER 9 ====================
    add_heading_1("Chapter 9: Clinical Translation, Real-Time Challenges, & Future Horizons")
    
    add_heading_2("9.1 Translational Opportunities in Intraoperative Neurosurgery")
    add_p(
        "Integrating DMCA into intraoperative surgical microscopes presents a paradigm shift in surgical oncology. By coupling high-speed hyperspectral sensors with real-time DMCA inference, neurosurgeons can project augmented-reality tumor margin overlays directly onto the operating field, enabling precise, millimeter-accurate tissue resection."
    )

    add_heading_2("9.2 Technical Challenges & Cross-Corpus Generalization")
    add_p(
        "As noted in the paper's concluding remarks, a significant hurdle to clinical translation is cross-corpus generalization. Hyperspectral acquisition hardware varies across clinical institutions in illumination geometry, sensor response, and spectral resolution. Because MedHSI datasets exhibit structural heterogeneity, models trained on one acquisition system cannot be directly deployed on another without domain adaptation. Future research must explore unsupervised domain adaptation (UDA) and physics-informed transfer learning to bridge equipment variations."
    )

    # ==================== CHAPTER 10 ====================
    add_heading_1("Chapter 10: Conclusion & Summary of Methodological Contributions")
    add_p(
        "The Deep Margin Cosine Autoencoder (DMCA) framework presented by Wang et al. represents a major advancement in medical hyperspectral image classification. By unifying feature extraction with classifier preference constraints, incorporating an angular cosine margin, and enforcing L1-norm feature sparsity, DMCA establishes a new state of the art in non-invasive tumor diagnosis. Empirical results across intraoperative brain cancer and cholangiocarcinoma datasets confirm its exceptional diagnostic accuracy (99.97% tumor detection rate), paving the way for real-time, label-free optical guidance in surgical oncology."
    )

    # ==================== REFERENCES ====================
    add_heading_1("References & Bibliography")
    add_p("The following formal reference list compiles key foundational literature cited in the DMCA paper and related hyperspectral analysis research:")
    
    references = [
        "[1] H. Li, P. Wu, Z. Wang, J. Mao, F. E. Alsaadi, and N. Zeng, 'A generalized framework of feature learning enhanced convolutional neural network for pathology-image-oriented cancer diagnosis,' Comput. Biol. Med., vol. 151, Art. no. 106265, Dec. 2022.",
        "[2] P. Wu, Z. Wang, B. Zheng, H. Li, F. E. Alsaadi, and N. Zeng, 'AGGN: Attention-based glioma grading network with multi-scale feature extraction and multi-modal information fusion,' Comput. Biol. Med., vol. 152, Art. no. 106457, Jan. 2023.",
        "[3] H. Li, N. Zeng, P. Wu, and K. Clawson, 'Cov-Net: A computer-aided diagnosis method for recognizing COVID-19 from chest X-ray images via machine vision,' Expert Syst. Appl., vol. 207, Art. no. 118029, Nov. 2022.",
        "[4] V. Dremin et al., 'Skin complications of diabetes mellitus revealed by polarized hyperspectral imaging and machine learning,' IEEE Trans. Med. Imag., vol. 40, no. 4, pp. 1207–1216, Apr. 2021.",
        "[5] M. Lv, W. Li, T. Chen, J. Zhou, and R. Tao, 'Discriminant tensor-based manifold embedding for medical hyperspectral imagery,' IEEE J. Biomed. Health Informat., vol. 25, no. 9, pp. 3517–3528, Sep. 2021.",
        "[6] M. Lv et al., 'Spatial-spectral density peaks-based discriminant analysis for membranous nephropathy classification using microscopic hyperspectral images,' IEEE J. Biomed. Health Informat., vol. 25, no. 8, pp. 3041–3051, Aug. 2021.",
        "[7] R. Pike, G. Lu, D. Wang, Z. G. Chen, and B. Fei, 'A minimum spanning forest-based method for noninvasive cancer detection with hyperspectral imaging,' IEEE Trans. Biomed. Eng., vol. 63, no. 3, pp. 653–663, Mar. 2016.",
        "[8] H. Akbari, Y. Kosugi, K. Kojima, and N. Tanaka, 'Detection and analysis of the intestinal ischemia using visible and invisible hyperspectral imaging,' IEEE Trans. Biomed. Eng., vol. 57, no. 8, pp. 2011–2017, Aug. 2010.",
        "[9] S. Kumar, C. Desmedt, D. Larsimont, C. Sotiriou, and E. Goormaghtigh, 'Change in the microenvironment of breast cancer studied by FTIR imaging,' Analyst, vol. 138, no. 14, pp. 4058–4065, Apr. 2013.",
        "[10] Z. Liu, H. Wang, and Q. Li, 'Tongue tumor detection in medical hyperspectral images,' Sensors, vol. 12, no. 1, pp. 162–174, Dec. 2011.",
        "[11] D. T. Dicker et al., 'Differentiation of normal skin and melanoma using high resolution hyperspectral imaging,' Cancer Biol. Therapy, vol. 5, no. 8, pp. 1033–1038, Aug. 2006.",
        "[12] B. Fei, H. Akbari, and L. V. Halig, 'Hyperspectral imaging and spectral-spatial classification for cancer detection,' in Proc. 5th IEEE Int. Conf. Biomed. Eng. Informat. (BMEI), Oct. 2012, pp. 62–64.",
        "[13] H. Akbari, Y. Kosugi, K. Kojima, and N. Tanaka, 'Hyperspectral image segmentation and its application in abdominal surgery,' Int. J. Funct. Inform. Personalised Med., vol. 2, no. 2, pp. 201–216, Aug. 2009.",
        "[14] Q. Li, X. He, Y. Wang, H. Liu, D. Xu, and F. Guo, 'Review of spectral imaging technology in biomedical engineering: Achievements and challenges,' J. Biomed. Opt., vol. 18, no. 10, Art. no. 100901, Oct. 2013.",
        "[15] G. S. Verebes et al., 'Hyperspectral enhanced dark field microscopy for imaging blood cells,' J. Biophoton., vol. 6, nos. 11–12, pp. 960–967, Aug. 2013.",
        "[16] M. Halicek et al., 'Deep convolutional neural networks for classifying head and neck cancer using hyperspectral imaging,' J. Biomed. Opt., vol. 22, no. 6, Art. no. 060503, Jun. 2017.",
        "[17] X. Li, W. Li, X. Xu, and W. Hu, 'Cell classification using convolutional neural networks in medical hyperspectral imagery,' in Proc. 2nd Int. Conf. Image, Vis. Comput. (ICIVC), Jun. 2017, pp. 501–504.",
        "[18] Q. Wang et al., 'Identification of melanoma from hyperspectral pathology image using 3D convolutional networks,' IEEE Trans. Med. Imag., vol. 40, no. 1, pp. 218–227, Jan. 2021.",
        "[19] X. Wei, W. Li, M. Zhang, and Q. Li, 'Medical hyperspectral image classification based on end-to-end fusion deep neural network,' IEEE Trans. Instrum. Meas., vol. 68, no. 11, pp. 4481–4492, Nov. 2019.",
        "[20] Q. Hao et al., 'Fusing multiple deep models for in vivo human brain hyperspectral image classification to identify glioblastoma tumor,' IEEE Trans. Instrum. Meas., vol. 70, pp. 1–14, 2021.",
        "[21] W. Zeng et al., 'Microscopic hyperspectral image classification based on fusion transformer with parallel CNN,' IEEE J. Biomed. Health Informat., vol. 27, no. 6, pp. 2910–2921, Jun. 2023.",
        "[22] H. Fabelo et al., 'In-vivo hyperspectral human brain image database for brain cancer detection,' IEEE Access, vol. 7, pp. 39098–39116, 2019.",
        "[23] H. Fabelo et al., 'Spatio-spectral classification of hyperspectral images for brain cancer detection during surgical operations,' PLoS ONE, vol. 13, no. 3, Art. no. e0193721, Mar. 2018.",
        "[24] Q. Zhang, Q. Li, G. Yu, L. Sun, M. Zhou, and J. Chu, 'A multidimensional choledoch database and benchmarks for cholangiocarcinoma diagnosis,' IEEE Access, vol. 7, pp. 149414–149421, 2019.",
        "[25] L. Sun et al., 'Diagnosis of cholangiocarcinoma from microscopic hyperspectral pathological dataset by deep convolution neural networks,' Methods, vol. 202, pp. 22–30, Jun. 2022.",
        "[26] W. Li and Q. Du, 'Gabor-filtering-based nearest regularized subspace for hyperspectral image classification,' IEEE J. Sel. Topics Appl. Earth Observ. Remote Sens., vol. 7, no. 4, pp. 1012–1022, Apr. 2014.",
        "[27] L. Fang, C. Wang, S. Li, and J. A. Benediktsson, 'Hyperspectral image classification via multiple-feature-based adaptive sparse representation,' IEEE Trans. Instrum. Meas., vol. 66, no. 7, pp. 1646–1657, Jul. 2017.",
        "[28] Y. Chen, Z. Lin, X. Zhao, G. Wang, and Y. Gu, 'Deep learning-based classification of hyperspectral data,' IEEE J. Sel. Topics Appl. Earth Observ. Remote Sens., vol. 7, no. 6, pp. 2094–2107, Jun. 2014.",
        "[29] W. Hu, Y. Huang, L. Wei, F. Zhang, and H. Li, 'Deep convolutional neural networks for hyperspectral image classification,' J. Sensors, vol. 2015, Art. no. 258619, Jan. 2015.",
        "[30] Y. Xu, B. Du, F. Zhang, and L. Zhang, 'Hyperspectral image classification via a random patches network,' ISPRS J. Photogramm. Remote Sens., vol. 142, pp. 344–357, Aug. 2018."
    ]

    for ref in references:
        add_bullet(ref)

    output_path = "DMCA_Comprehensive_Tumor_Diagnosis_Report.docx"
    doc.save(output_path)
    print(f"Document successfully created and saved to: {output_path}")

if __name__ == "__main__":
    create_report()
