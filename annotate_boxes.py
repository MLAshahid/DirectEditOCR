import json
from pathlib import Path
from PIL import Image, ImageTk
import tkinter as tk

HELP = """Controls:
- Drag Left Mouse: draw a box
- Right Click: select a box
- Z: undo last   |  D: delete selected
- N / P: next / previous page
- S: save JSON   |  Q / Esc: quit
- + / - : zoom in / out   |  Mouse Wheel: scroll (Shift+Wheel: horizontal)
"""

class Annotator:
    def __init__(self, images, out_json, init_zoom=1.5):
        self.images = images
        self.out_json = out_json
        self.index = 0
        self.zoom = float(init_zoom)

        self.data = [{
            "path": str(p),
            "width": Image.open(p).size[0],
            "height": Image.open(p).size[1],
            "boxes": []
        } for p in images]

        self.root = tk.Tk()
        self.root.title("Box Annotator")
        self.root.geometry("1300x900")

        # Canvas + scrollbars
        self.canvas = tk.Canvas(self.root, bg="black", highlightthickness=0)
        self.hbar = tk.Scrollbar(self.root, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.vbar = tk.Scrollbar(self.root, orient=tk.VERTICAL,   command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=self.hbar.set, yscrollcommand=self.vbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.hbar.grid(row=1, column=0, sticky="ew")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        self.status = tk.Label(self.root, text=HELP, anchor="w", justify="left")
        self.status.grid(row=2, column=0, columnspan=2, sticky="ew")

        self.scale = 1.0
        self.tkimg = None
        self.img = None

        self.rect_start = None
        self.rect_id = None
        self.selected_idx = None

        # events
        self.root.bind("<Configure>", self._on_resize)
        self.canvas.bind("<Button-1>", self._on_down)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_up)
        self.canvas.bind("<Button-3>", self._on_right_click)

        # mouse wheel scroll
        self.canvas.bind("<MouseWheel>", self._on_wheel)          # Win
        self.canvas.bind("<Shift-MouseWheel>", self._on_wheel_h)

        # zoom keys
        for key in ("<KeyPress-plus>", "<KeyPress-KP_Add>", "<KeyPress-equal>"):
            self.root.bind(key, lambda e: self._zoom(1.15))
        for key in ("<KeyPress-minus>", "<KeyPress-KP_Subtract>"):
            self.root.bind(key, lambda e: self._zoom(1/1.15))

        self.root.bind("<KeyPress-n>", lambda e: self.next())
        self.root.bind("<KeyPress-p>", lambda e: self.prev())
        self.root.bind("<KeyPress-s>", lambda e: self.save())
        self.root.bind("<KeyPress-z>", lambda e: self.undo())
        self.root.bind("<KeyPress-d>", lambda e: self.delete_selected())
        self.root.bind("<Escape>",    lambda e: self.quit())
        self.root.bind("<KeyPress-q>",lambda e: self.quit())

        self._load_image()
        self.root.mainloop()

    def _on_wheel(self, e):
        self.canvas.yview_scroll(int(-1*(e.delta/120)), "units")

    def _on_wheel_h(self, e):
        self.canvas.xview_scroll(int(-1*(e.delta/120)), "units")

    def _zoom(self, factor):
        self.zoom = max(0.2, min(6.0, self.zoom * factor))
        self._render()

    def _on_resize(self, event):
        # keep best-fit base, then apply user zoom
        self._render()

    def _load_image(self):
        img_path = self.images[self.index]
        self.img = Image.open(img_path).convert("RGB")
        self._render()

    def _render(self):
        if not self.img:
            return

        cw = max(1, self.canvas.winfo_width())
        ch = max(1, self.canvas.winfo_height())
        iw, ih = self.img.size

        best_fit = min(cw / iw, ch / ih)
        self.scale = max(0.001, best_fit * self.zoom)

        tw = max(1, int(iw * self.scale))
        th = max(1, int(ih * self.scale))

        disp = self.img.resize((tw, th))
        self.tkimg = ImageTk.PhotoImage(disp)

        self.canvas.delete("all")
        self.canvas.create_image(0, 0, image=self.tkimg, anchor="nw")
        self.canvas.config(scrollregion=(0, 0, tw, th))

        # draw boxes
        for i, b in enumerate(self.data[self.index]["boxes"]):
            x0 = int(b["left"] * self.scale)
            y0 = int(b["top"] * self.scale)
            x1 = int((b["left"] + b["width"]) * self.scale)
            y1 = int((b["top"] + b["height"]) * self.scale)
            color = "yellow" if i == self.selected_idx else "cyan"
            self.canvas.create_rectangle(x0, y0, x1, y1, outline=color, width=2)
            self.canvas.create_text(x0 + 4, y0 + 10, text=str(i + 1), anchor="w", fill=color)

        self.status.config(
            text=f"{HELP}\nImage {self.index+1}/{len(self.images)}: "
                 f"{self.images[self.index]}  Boxes: {len(self.data[self.index]['boxes'])}  Zoom: {self.zoom:.2f}x"
        )

    def _canvas_xy(self, e):
        return (self.canvas.canvasx(e.x), self.canvas.canvasy(e.y))

    def _on_down(self, e):
        cx, cy = self._canvas_xy(e)
        self.selected_idx = self._hit_test(cx, cy)
        if self.selected_idx is not None:
            self._render(); return
        self.rect_start = (cx, cy)
        if self.rect_id:
            self.canvas.delete(self.rect_id); self.rect_id = None

    def _on_drag(self, e):
        if not self.rect_start: return
        cx, cy = self._canvas_xy(e)
        x0, y0 = self.rect_start
        if self.rect_id:
            self.canvas.coords(self.rect_id, x0, y0, cx, cy)
        else:
            self.rect_id = self.canvas.create_rectangle(x0, y0, cx, cy, outline="red", width=2)

    def _on_up(self, e):
        if not self.rect_start: return
        cx, cy = self._canvas_xy(e)
        x0, y0 = self.rect_start
        self.rect_start = None

        if abs(cx - x0) < 5 or abs(cy - y0) < 5:
            if self.rect_id:
                self.canvas.delete(self.rect_id); self.rect_id = None
            return

        left   = int(min(x0, cx) / self.scale)
        top    = int(min(y0, cy) / self.scale)
        width  = int(abs(cx - x0) / self.scale)
        height = int(abs(cy - y0) / self.scale)

        self.data[self.index]["boxes"].append({"left": left, "top": top, "width": width, "height": height})
        if self.rect_id:
            self.canvas.delete(self.rect_id); self.rect_id = None
        self._render()

    def _on_right_click(self, e):
        cx, cy = self._canvas_xy(e)
        self.selected_idx = self._hit_test(cx, cy)
        self._render()

    def _hit_test(self, x, y):
        for i, b in enumerate(self.data[self.index]["boxes"]):
            x0 = b["left"] * self.scale; y0 = b["top"] * self.scale
            x1 = (b["left"] + b["width"]) * self.scale
            y1 = (b["top"]  + b["height"]) * self.scale
            if x0 <= x <= x1 and y0 <= y <= y1:
                return i
        return None

    def undo(self):
        boxes = self.data[self.index]["boxes"]
        if boxes: boxes.pop(); self._render()

    def delete_selected(self):
        i = self.selected_idx
        if i is not None and 0 <= i < len(self.data[self.index]["boxes"]):
            self.data[self.index]["boxes"].pop(i); self.selected_idx = None; self._render()

    def next(self):
        if self.index < len(self.images) - 1:
            self.index += 1; self.selected_idx = None; self._load_image()

    def prev(self):
        if self.index > 0:
            self.index -= 1; self.selected_idx = None; self._load_image()

    def save(self):
        Path(self.out_json).write_text(json.dumps({"images": self.data}, indent=2), encoding="utf-8")
        print("Saved", self.out_json)

    def quit(self):
        self.save(); self.root.destroy()


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--images", required=True)
    ap.add_argument("--out", default="boxes.json")
    ap.add_argument("--zoom", type=float, default=1.5, help="Initial zoom multiplier (default 1.5)")
    args = ap.parse_args()

    p = Path(args.images)
    if p.is_dir():
        imgs = []
        for ext in ("*.png","*.jpg","*.jpeg","*.tif","*.tiff","*.bmp","*.webp"):
            imgs.extend(sorted(p.glob(ext)))
        if not imgs: raise SystemExit("No images found in folder.")
    else:
        imgs = [p]

    Annotator([str(x) for x in imgs], args.out, init_zoom=args.zoom)

if __name__ == "__main__":
    main()
