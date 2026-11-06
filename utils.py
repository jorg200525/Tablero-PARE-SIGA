"""
Módulo de utilidades comunes para tableros Streamlit.
Contiene funciones reutilizables: detección de IP, banner de compartir, etc.

Uso:
    from utils import mostrar_ip_sidebar, mostrar_ip_header, obtener_ip_local
"""

import socket
import streamlit as st


# ============================================================
# DETECCIÓN DE IP
# ============================================================
def obtener_ip_local():
    """
    Detecta la IP local de la PC en la red.
    
    Usa el truco de conectarse a 8.8.8.8 (sin enviar datos reales) para
    que el sistema operativo nos diga qué interfaz está activa. Esto evita
    devolver 127.0.0.1 o IPs virtuales cuando hay VPN.
    
    Returns:
        str: IP local (ej: '192.168.1.45') o 'localhost' si falla.
    """
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "localhost"


# ============================================================
# MOSTRAR IP EN EL SIDEBAR
# ============================================================
def mostrar_ip_sidebar(st, puerto=8501, estilo="caja", titulo="Compartir en red"):
    """
    Muestra la URL de acceso al tablero en el sidebar.
    
    Args:
        st: El módulo streamlit ya importado.
        puerto (int): Puerto donde corre Streamlit (default: 8501).
        estilo (str): "caja" (con fondo azul), "simple" (texto plano),
                      "codigo" (bloque de código).
        titulo (str): Texto que aparece antes de la URL.
    
    Ejemplo:
        mostrar_ip_sidebar(st, puerto=8501, estilo="caja")
    """
    url = f"http://{obtener_ip_local()}:{puerto}"
    
    if estilo == "caja":
        st.sidebar.markdown(f"""
<div style="background-color: #e8f4f8; padding: 12px; border-radius: 8px;
            border-left: 4px solid #4A90E2; margin-bottom: 15px;">
    <small style="color: #666;">🌐 <b>{titulo}:</b></small><br>
    <code style="font-size: 13px; color: #1f6feb;">{url}</code>
</div>
""", unsafe_allow_html=True)
    
    elif estilo == "codigo":
        st.sidebar.markdown(f"🌐 **{titulo}:**")
        st.sidebar.code(url, language="text")
    
    else:  # simple
        st.sidebar.markdown(f"🌐 **{titulo}:** `{url}`")


# ============================================================
# MOSTRAR IP EN EL HEADER (arriba del contenido)
# ============================================================
def mostrar_ip_header(st, puerto=8501, titulo="Comparte este tablero"):
    """
    Muestra la URL de acceso como un banner azul arriba del contenido principal.
    
    Ejemplo:
        mostrar_ip_header(st, puerto=8501)
    """
    url = f"http://{obtener_ip_local()}:{puerto}"
    st.markdown(f"""
<div style="background-color: #e8f4f8; padding: 10px 15px; border-radius: 8px;
            border-left: 4px solid #4A90E2; margin-bottom: 15px;">
    🌐 <b>{titulo}:</b> <code style="font-size: 14px;">{url}</code>
    <span style="color: #666; font-size: 12px;">(misma red WiFi/LAN)</span>
</div>
""", unsafe_allow_html=True)


# ============================================================
# MOSTRAR QR (opcional, requiere: pip install qrcode[pil])
# ============================================================
def mostrar_qr_sidebar(st, puerto=8501):
    """
    Muestra un QR en el sidebar con la URL del tablero.
    Los compañeros pueden escanearlo con el celular para abrir el tablero.
    
    Requiere: pip install qrcode[pil]
    
    Ejemplo:
        mostrar_qr_sidebar(st, puerto=8501)
    """
    try:
        import qrcode
        from io import BytesIO
        
        url = f"http://{obtener_ip_local()}:{puerto}"
        qr = qrcode.QRCode(box_size=4, border=2)
        qr.add_data(url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")
        
        buf = BytesIO()
        img.save(buf, format="PNG")
        
        st.sidebar.markdown("##### 📱 Escanea para abrir")
        st.sidebar.image(buf.getvalue(), width=180)
        st.sidebar.caption(url)
    except ImportError:
        st.sidebar.warning("⚠️ Instala `qrcode[pil]` para usar esta función.")