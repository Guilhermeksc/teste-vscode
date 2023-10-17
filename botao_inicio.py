from botoes.config_botao import carregar_imagem_redimensionada
from diretorios import ICONS_DIR

# Mantendo uma referência global à imagem
global_ref_image = None

def reset_canvas_to_default(canvas):
    global global_ref_image
    for widget in canvas.winfo_children():
        widget.destroy()
    canvas.create_rectangle(5, 5, 925, 675, fill="#FFFFFF", outline="#FFFFFF", width=5)  # Cor alterada para branco
    global_ref_image = carregar_imagem_redimensionada(ICONS_DIR / "ceimbra.png")
    image_id = canvas.create_image(470, 350, image=global_ref_image)
    canvas.create_text(470, 90, text="CeIMBra", font=("Arial", 72, "bold"))
