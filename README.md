<img width="100" height="100" alt="image" src="https://github.com/user-attachments/assets/d3ae1368-fb8e-4a44-8958-b73ee19d7eeb" />

# PowerPoint Icon Builder
A PowerPoint VSTO add-in that provides options to build and export .ico icons.

#### Supported sizes
The .ico file will be encoded for the 256x256 size as a png, and for the remaining sizes it will be encoded as a 32-bit bitmap.

The full list of supported sizes for the Icon Builder extension are:
- 256x256
- 64x64
- 48x48
- 40x40
- 32x32
- 16x16
- 20x20
- 24x24

### Right Click Menu
<img width="285" height="186" alt="image" src="https://github.com/user-attachments/assets/9e180e03-5354-4a52-832e-39215e47e5d9" />

Select any image, object, or group, and right click for the new option "Save as Icon (.ico)...". This option will automatically scale the icon to all supported sizes and export it to your disk as a .ico file.

### Ribbon
<img width="837" height="184" alt="image" src="https://github.com/user-attachments/assets/c4fea168-f1c4-4ed9-b535-5b26c68e2d1f" />

The ribbon features these functions:

- **Save Selection as Icon**: Save the selected picture, shape, or group as a .ico file. The ico file will automatically save with all the sizes supported by the Icon Builder extension.
- **Icon Editor Pane**: This toggles open or close on the [Icon Editing Pane](#icon-editing-pane)
- **Guide Sizes**: A list of checkboxes that allows the user to determine which guide sizes they would like to insert on the page. See [Guide Squares](#guide-squares)
- **Insert Guide Squares**: Inserts guide squares on the center of the PowerPoint. See [Guide Squares](#guide-squares)

### Guide Squares
<img width="889" height="463" alt="image" src="https://github.com/user-attachments/assets/20f18654-f63f-4c4a-b209-7e59835b75da" />

Good icons should look different at different sizes. Smaller icons should be simplified, so that they look distinctive even with few pixels to work with. Larger icons can be more detailed and colorful.

To help design simplified icons at smaller sizes, you can create guide squares that give you boxes of each relative size supported by the Icon Builder extension. You can design your icon, resize them to fit in each box, and judge how you would like to modify or simplify each size based on the guide squares. Once you are happy with how your design looks, you can send the different sizes of your icon to the [Icon Editing Pane](#icon-editing-pane)

### Icon Editing Pane
<img width="402" height="970" alt="image" src="https://github.com/user-attachments/assets/f17c3d60-926a-4835-8e3b-f74b2df75cf9" />

The Icon editing pane allows you to assign an icon to each individual size, so that you can have different icons for different resolutions.

Select an image, object, or group in PowerPoint and click "Assign" in the Icon Editing Pane to assign it to a specific size. The selection in PowerPoint will be automatically scaled to whichever size you pick once you assign it to that size. If you do not want to use a given size in your ICO file, you can simply not assign anything to that size.

If you'd like to replace the icon for a certain size, you can select a different image, object, or group at click "assign" for that size. Once you are done populating all the different resolutions your icon can be rendered at, you can press the green "Export as Icon (.ico)..." at the bottom of the pane in order to export your icon.

### Preview Icon
<img width="995" height="676" alt="image" src="https://github.com/user-attachments/assets/32173cb6-e069-45cf-8f5b-858b745fb455" />

Before the icon can be saved, you have one last chance to preview the icon at every resolution. Once you are satisfied with how the icon works, you can hit "Save" and you will be prompted for a location on your computer to save the icon to.

## AI Disclosure
This project has been entirely vibe coded and minimally reviewed. There's no SLA or warranty of any kind, not recommended for production setups.
