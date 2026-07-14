"""
Step 1: Crop architecture diagrams from friend's slides and my slides
Step 2: Build exactly 40 slides with TYPED text + diagram images
"""
from PIL import Image
import os

BASE = r"C:\Users\vujja\OneDrive\Documents\OneDrive\Desktop\btpdataset"
IMG = os.path.join(BASE, "_temp_scripts", "slide_images_work", "slide_images")
FRD = os.path.join(BASE, "_temp_scripts", "slide_images_work", "friend_slides")
CROP = os.path.join(BASE, "_temp_scripts", "cropped_diagrams")
os.makedirs(CROP, exist_ok=True)

def crop_img(src, dst, box):
    """box = (left, upper, right, lower) as fractions of image size"""
    if not os.path.exists(src):
        print(f"  SKIP: {src}")
        return
    img = Image.open(src)
    w, h = img.size
    crop_box = (int(box[0]*w), int(box[1]*h), int(box[2]*w), int(box[3]*h))
    cropped = img.crop(crop_box)
    cropped.save(dst)
    print(f"  Cropped: {os.path.basename(dst)} ({cropped.size[0]}x{cropped.size[1]})")

# Friend's slides - crop architecture diagrams
# Slide 10: SRCNN butterfly/pipeline diagram (right half)
crop_img(f"{FRD}/slide_10_rot.png", f"{CROP}/srcnn_arch.png", (0.5, 0.05, 1.0, 0.75))
# Slide 21: SRGAN workflow diagrams (right half)
crop_img(f"{FRD}/slide_21_rot.png", f"{CROP}/srgan_arch.png", (0.45, 0.03, 1.0, 0.85))
# Slide 26: Vision Transformer diagram (right half)
crop_img(f"{FRD}/slide_26_rot.png", f"{CROP}/swinir_arch.png", (0.45, 0.3, 1.0, 0.95))
# Slide 29: SSM diagram (bottom center)
crop_img(f"{FRD}/slide_29_rot.png", f"{CROP}/mamba_ssm.png", (0.25, 0.45, 0.95, 0.95))
# Slide 7: Evolution timeline (right half with dots)
crop_img(f"{FRD}/slide_07_rot.png", f"{CROP}/evolution.png", (0.35, 0.1, 0.95, 0.95))

# My slides - crop just the diagram areas
for i in range(2, 29):
    src = f"{IMG}/slide_{i:02d}.png"
    dst = f"{CROP}/my_{i:02d}.png"
    if os.path.exists(src):
        # My slides are landscape, crop out any excessive borders
        crop_img(src, dst, (0.02, 0.05, 0.98, 0.95))

print("\nAll crops done!")
