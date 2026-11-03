import re
from supabase import create_client, Client

# ==========================================
# CREDENCIALES DE SUPABASE
# ==========================================
url: str = "https://qkduhwrmwvtpptxwumea.supabase.co"
key: str = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFrZHVod3Jtd3Z0cHB0eHd1bWVhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTEzODc3NjcsImV4cCI6MjEwNjk2Mzc2N30.gnImdIvu_7AXGqKR7_j1gIi3HLqhwUPOLfblZy-n2bU"
supabase: Client = create_client(url, key)

ID_EDICION = 1 

def limpiar_e_inyectar():
    print("1. Leyendo el HTML...")
    with open('texto_libro.html', 'r', encoding='utf-8') as archivo_html:
        html_crudo = archivo_html.read()
        
    print("2. Limpiando la basura de Google AI ()...")
    # Esta línea mágica busca cualquier etiqueta y la elimina
    html_limpio = re.sub(r'\[cite:[\d\s,]+\]', '', html_crudo)
    
    print("3. Subiendo el HTML inmaculado a Supabase...")
    supabase.table("ediciones").update({
        "contenido_presentacion": html_limpio
    }).eq("id", ID_EDICION).execute()
    
    print("✅ ¡ÉXITO! El libro está limpio y cargado en el sistema.")

if __name__ == "__main__":
    limpiar_e_inyectar()