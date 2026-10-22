import os
import tempfile
from datetime import datetime
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_word_doc(transcript_text, filename):
    """ฟังก์ชันสำหรับสร้างไฟล์เอกสาร Word"""
    doc = Document()
    
    # ตั้งค่าฟอนต์
    style = doc.styles['Normal']
    font = style.font
    font.name = 'TH Sarabun New'
    font.size = Pt(16)

    # สร้างหัวกระดาษ
    title = doc.add_heading('บันทึกการถอดเทปการประชุม', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # สร้างรายละเอียดไฟล์
    meta = doc.add_paragraph()
    meta.add_run(f"ไฟล์ต้นฉบับ: {filename}\n").bold = True
    meta.add_run(f"วันที่ประมวลผล: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n").bold = True
    doc.add_paragraph("-" * 60)

    # ใส่เนื้อหาที่ถอดเสียงมา
    doc.add_heading('เนื้อหาการสนทนา', level=2)
    doc.add_paragraph(transcript_text)

    # บันทึกลงไฟล์ชั่วคราวและส่ง Path คืนกลับไป
    temp_docx = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(temp_docx.name)
    return temp_docx.name

# ==========================================
# ส่วนของการสร้างหน้าเว็บ (Streamlit UI)
# ==========================================
st.set_page_config(page_title="AI Audio to Word", page_icon="🎙️", layout="centered")

st.title("🎙️ ระบบถอดไฟล์เสียงเป็น Word ด้วย AI")
st.markdown("ระบบนี้ใช้ **Google Gemini 1.5 Flash** ในการประมวลผล รองรับไฟล์เสียงขนาดใหญ่สูงสุด 2 GB")

# แถบด้านข้างสำหรับตั้งค่า
with st.sidebar:
    st.header("⚙️ การตั้งค่าระบบ")
    api_key = st.text_input("ใส่ Google Gemini API Key", type="password", help="รับฟรีได้ที่ https://aistudio.google.com/")
    if api_key:
        st.success("รับ API Key แล้ว")
    
    st.markdown("---")
    st.markdown("### ตัวเลือกเสริม")
    do_summary = st.checkbox("ให้ AI สรุปวาระการประชุมและ Action Items ต่อท้ายเอกสาร", value=True)

# ส่วนอัปโหลดไฟล์หน้าหลัก
uploaded_file = st.file_uploader("เลือกไฟล์เสียงที่ต้องการถอด (รองรับ .mp3, .wav, .m4a, .aac)", type=['mp3', 'wav', 'm4a', 'aac'])

if uploaded_file is not None:
    if not api_key:
        st.warning("⚠️ กรุณาใส่ API Key ที่แถบด้านข้างก่อนเริ่มทำงาน")
    else:
        if st.button("🚀 เริ่มถอดเสียงและสร้างเอกสาร", use_container_width=True):
            
            # ตั้งค่า API
            genai.configure(api_key=api_key)
            
            try:
                # 1. บันทึกไฟล์ลง Temp File เพื่อเตรียมอัปโหลด
                with st.spinner("⏳ ขั้นตอนที่ 1/4: กำลังเตรียมไฟล์..."):
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmp_audio:
                        tmp_audio.write(uploaded_file.getvalue())
                        tmp_audio_path = tmp_audio.name

                # 2. อัปโหลดไฟล์ขึ้น Google Cloud (Gemini Files API)
                with st.spinner("☁️ ขั้นตอนที่ 2/4: กำลังอัปโหลดไฟล์เข้าระบบ AI..."):
                    audio_file = genai.upload_file(path=tmp_audio_path)
                
                # 3. สั่งให้ AI ถอดเสียง
                with st.spinner("🧠 ขั้นตอนที่ 3/4: AI กำลังฟังและถอดเสียง (อาจใช้เวลา 1-3 นาที)..."):
                    model = genai.GenerativeModel('models/gemini-1.5-flash')
                    
                    if do_summary:
                        prompt = "ถอดเสียงจากไฟล์นี้ออกมาเป็นข้อความภาษาไทยให้ถูกต้องและครบถ้วนที่สุด และในตอนท้ายของข้อความ ให้สรุปประเด็นสำคัญ และสิ่งที่ต้องทำต่อ (Action Items) จัดหน้าให้สวยงาม"
                    else:
                        prompt = "ถอดเสียงจากไฟล์นี้ออกมาเป็นภาษาไทยให้ถูกต้องและครบถ้วนที่สุด พิมพ์เฉพาะข้อความที่ได้ยินเท่านั้น"

                    response = model.generate_content([prompt, audio_file])
                    transcript_result = response.text
                    
                # 4. สร้างไฟล์ Word
                with st.spinner("📄 ขั้นตอนที่ 4/4: กำลังสร้างเอกสาร Word..."):
                    docx_path = create_word_doc(transcript_result, uploaded_file.name)
                    
                # 5. ทำความสะอาดไฟล์จากระบบ
                genai.delete_file(audio_file.name) # ลบไฟล์ออกจาก Google
                os.remove(tmp_audio_path) # ลบไฟล์ออกจากเซิร์ฟเวอร์เรา
                
                st.success("✅ ประมวลผลเสร็จสมบูรณ์!")
                
                # ปุ่มดาวน์โหลดไฟล์ Word
                with open(docx_path, "rb") as file:
                    st.download_button(
                        label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ Word (.docx)",
                        data=file,
                        file_name=f"Transcript_{uploaded_file.name}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
            except Exception as e:
                st.error(f"❌ เกิดข้อผิดพลาด: {e}")
                st.info("คำแนะนำ: โปรดตรวจสอบว่า API Key ถูกต้องและไฟล์เสียงไม่ได้ชำรุด")