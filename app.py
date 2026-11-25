import os
import re
import tempfile
import streamlit as st
import whisper
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


st.set_page_config(
    page_title="MP3 語音轉排版 Word 工具",
    page_icon="📝",
    layout="centered",
)

st.title("📝 MP3 語音轉排版 Word 工具")
st.caption("上傳語音檔後自動轉錄為純文字並排版，直接下載 Word（.docx）檔案。")


@st.cache_resource(show_spinner="正在載入 Whisper 語音模型，首次啟動需稍候...")
def load_whisper_model(model_size: str):
    return whisper.load_model(model_size)


def format_text_into_paragraphs(text: str, sentences_per_para: int = 4) -> list[str]:
    """
    將連續文字清理多餘空白，並依句子結束符號（。！？!?）重新組合成較易閱讀的段落。
    """
    cleaned = re.sub(r"\s+", " ", text).strip()
    if not cleaned:
        return []

    # 若辨識結果完全沒有句尾標點，則按字數切段
    if not re.search(r"[。！？!?]", cleaned):
        chunk_size = 120
        return [
            cleaned[i : i + chunk_size].strip()
            for i in range(0, len(cleaned), chunk_size)
            if cleaned[i : i + chunk_size].strip()
        ]

    sentences = re.findall(r"[^。！？!?]+[。！？!?]*", cleaned)
    sentences = [s.strip() for s in sentences if s.strip()]

    paragraphs = []
    current_para = []

    for sentence in sentences:
        current_para.append(sentence)
        if len(current_para) >= sentences_per_para:
            paragraphs.append("".join(current_para))
            current_para = []

    if current_para:
        paragraphs.append("".join(current_para))

    return paragraphs


def build_docx_bytes(audio_filename: str, paragraphs: list[str]) -> bytes:
    """
    生成排版的 Word (.docx) 並回傳位元組資料供 st.download_button 下載。
    """
    base_name = os.path.splitext(audio_filename)[0]
    doc = Document()

    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    normal_style = doc.styles["Normal"]
    normal_style.font.name = "Microsoft JhengHei"
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(34, 34, 34)

    title_para = doc.add_heading(f"語音轉文字內容 — {base_name}", level=1)
    title_para.alignment = WD_ALIGN_PARAGRAPH.LEFT

    subtitle = doc.add_paragraph(f"來源檔案：{audio_filename}")
    subtitle.runs[0].font.size = Pt(9.5)
    subtitle.runs[0].font.color.rgb = RGBColor(120, 120, 120)
    subtitle.paragraph_format.space_after = Pt(14)

    for para_text in paragraphs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.first_line_indent = Pt(24)

        run = p.add_run(para_text)
        run.font.name = "Microsoft JhengHei"
        run.font.size = Pt(12)

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp_docx:
        tmp_docx_path = tmp_docx.name

    try:
        doc.save(tmp_docx_path)
        with open(tmp_docx_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(tmp_docx_path):
            os.remove(tmp_docx_path)


# --- 側邊欄設定 ---
with st.sidebar:
    st.header("轉錄與排版設定")
    # 免費雲端建議使用 base 或 small；medium/large 記憶體需求高，可能導致容器重啟
    model_size = st.selectbox(
        "Whisper 模型大小",
        options=["base", "small", "medium"],
        index=1,
        help="模型越大準確率越高，但免費雲端執行時間與記憶體用量也越高。",
    )
    sentences_per_para = st.slider(
        "每段句數",
        min_value=2,
        max_value=8,
        value=4,
        step=1,
    )
    prompt_hint = st.text_input(
        "語音提示詞（引導繁體與標點）",
        value="以下是繁體中文語音內容，請使用正確的繁體中文與標點符號。",
    )

# --- 主畫面：上傳與轉換 ---
uploaded_file = st.file_uploader(
    "請選擇要轉換的 MP3 語音檔",
    type=["mp3"],
)

if uploaded_file is not None:
    st.audio(uploaded_file, format="audio/mp3")

    if st.button("開始轉換並生成 Word", type="primary", use_container_width=True):
        with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmp_audio:
            tmp_audio.write(uploaded_file.getvalue())
            tmp_audio_path = tmp_audio.name

        try:
            model = load_whisper_model(model_size)

            with st.spinner("正在辨識語音並整理段落，請稍候..."):
                result = model.transcribe(
                    tmp_audio_path,
                    verbose=False,
                    initial_prompt=prompt_hint.strip() or None,
                )

            raw_text = result.get("text", "").strip()
            if not raw_text:
                st.error("未能辨識出任何文字內容，請確認語音檔是否有聲音。")
            else:
                paragraphs = format_text_into_paragraphs(
                    raw_text,
                    sentences_per_para=sentences_per_para,
                )
                docx_bytes = build_docx_bytes(uploaded_file.name, paragraphs)
                base_name = os.path.splitext(uploaded_file.name)[0]
                download_filename = f"{base_name}_逐字稿排版.docx"

                st.success("轉換與排版完成！")

                with st.expander("預覽整理後的文字內容", expanded=True):
                    st.write("\n\n".join(paragraphs))

                st.download_button(
                    label="⬇️ 下載排版的 Word 檔（.docx）",
                    data=docx_bytes,
                    file_name=download_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
        finally:
            if os.path.exists(tmp_audio_path):
                os.remove(tmp_audio_path)
