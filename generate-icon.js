const sharp = require('sharp');
const fs = require('fs');
const path = require('path');

async function generateIcon() {
  const sizes = [16, 24, 32, 48, 64, 128, 256];
  const pngPath = path.join(__dirname, 'build', 'icon.png');
  const icoPath = path.join(__dirname, 'build', 'icon.ico');

  // 使用 png-to-ico 库生成 ICO
  const pngToIco = require('png-to-ico');

  // 为每个尺寸生成 PNG
  const pngBuffers = await Promise.all(
    sizes.map(size =>
      sharp(pngPath)
        .resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
        .png()
        .toBuffer()
    )
  );

  // 生成 ICO
  const icoBuffer = await pngToIco(pngBuffers);
  fs.writeFileSync(icoPath, icoBuffer);
  console.log('Icon generated successfully:', icoPath);
}

generateIcon().catch(console.error);
