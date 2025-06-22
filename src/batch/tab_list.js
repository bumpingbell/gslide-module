// Tab list module for Google Slides
/**
 * 主流程：先刪除所有舊的 TAB_LIST_* 元件，再批次建立新的頁籤列，並加入分頁
 */
function processTabs(slides) {
  const CONFIG = {
    totalWidth: 720,
    height: 14, // Initial height, will be dynamically adjusted
    y: 0,
    fontSize: 8,
    padding: 10, // Increased padding for better appearance and to accommodate longer text
    spacing: 2, // Spacing between tabs
    mainColor: main_color,
    mainFont: main_font_family,
    bgColor: '#FFFFFF',
    inactiveTextColor: '#888888',
    minWidth: 50,
    maxTabHeight: 40, // New: Maximum height a tab can expand to for multi-line text
    lineHeightFactor: 1.2, // New: Factor to determine height per line
    maxLines: 2 // New: Maximum number of lines a tab heading can have
  };

  const sections = getSectionHeaders(slides);
  if (sections.length === 0) return;

  // First, delete old tabs client-side
  slides.forEach((slide, idx) => {
    if (idx === 0) return; // Skip cover slide
    deleteOldTabsClient(slide);
  });

  const requests = [];
  let currentSectionIdx = -1;
  const totalPages = slides.length;

  // Calculate estimated character width based on font size. This is an approximation.
  // We'll iterate to find max width needed and then determine if multi-line is necessary.
  const estCharW = CONFIG.fontSize * 0.6; // Adjusted for better fit

  // Pre-calculate desired widths and heights for each tab
  const tabData = sections.map(sec => {
    const text = sec.title;
    // Calculate ideal width based on text length and padding
    let idealWidth = text.length * estCharW + CONFIG.padding * 2;
    idealWidth = Math.max(idealWidth, CONFIG.minWidth);

    // Determine if text needs wrapping and calculate height
    let lines = 1;
    let actualWidth = idealWidth; // Start with ideal width
    if (idealWidth > CONFIG.totalWidth / sections.length && sections.length > 1) {
      // If a single tab is too wide relative to available space per tab, consider wrapping
      // This is a heuristic. More precise calculation would involve text breaking algorithms.
      const maxAllowedWidthPerTab = (CONFIG.totalWidth - (sections.length - 1) * CONFIG.spacing) / sections.length;
      if (idealWidth > maxAllowedWidthPerTab) {
        // Estimate how many lines if we force width to maxAllowedWidthPerTab
        lines = Math.min(Math.ceil((text.length * estCharW) / (maxAllowedWidthPerTab - CONFIG.padding * 2)), CONFIG.maxLines);
        actualWidth = Math.max(maxAllowedWidthPerTab, CONFIG.minWidth); // Use max allowed width or minWidth
      }
    }
    
    // Ensure width is not excessively small if text is short and lines increased
    if (lines > 1 && actualWidth < CONFIG.minWidth * 1.5) { // If multiline, ensure a slightly larger minimum width
      actualWidth = CONFIG.minWidth * 1.5;
    }


    let calculatedHeight = CONFIG.height;
    if (lines > 1) {
      calculatedHeight = Math.min(CONFIG.height * CONFIG.lineHeightFactor * lines, CONFIG.maxTabHeight);
    }
    
    // Ensure minimum height even for single line to maintain consistency
    calculatedHeight = Math.max(calculatedHeight, CONFIG.height);


    return {
      title: text,
      index: sec.index,
      slideId: sec.slideId,
      width: actualWidth,
      height: calculatedHeight,
      lines: lines // Store number of lines calculated
    };
  });

  // Now, adjust all tab heights to the maximum calculated height among them to keep them uniform
  const uniformTabHeight = tabData.reduce((maxH, tab) => Math.max(maxH, tab.height), CONFIG.height);
  // Also adjust total width based on the newly calculated widths
  const totalTabsWidth = tabData.reduce((sum, tab) => sum + tab.width, 0) + CONFIG.spacing * (tabData.length - 1);

  // Calculate the starting X position to center the tabs
  const xStart = Math.max((CONFIG.totalWidth - totalTabsWidth) / 2, 0);

  slides.forEach((slide, idx) => {
    if (idx === 0) return;
    const slideId = slide.getObjectId();

    // Determine current section index
    if (currentSectionIdx + 1 < sections.length && idx >= sections[currentSectionIdx + 1].index) {
      currentSectionIdx++;
    }
    appendPageNumberToSlide({ slideId, requests, currentPage: idx + 1, totalPages, config: CONFIG });
    if (slide.getLayout().getLayoutName() === 'SECTION_HEADER') return;

    const currentSection = currentSectionIdx >= 0 ? currentSectionIdx : -1;
    // Add tab list
    appendTabListToSlide({
      slideId,
      requests,
      sections: tabData, // Pass the processed tabData
      currentSection,
      config: { ...CONFIG, height: uniformTabHeight, xStart: xStart } // Pass updated height and xStart
    });
  });

  if (requests.length) {
    Slides.Presentations.batchUpdate({ requests }, SlidesApp.getActivePresentation().getId());
  }
}

/** 讀出所有 SECTION_HEADER 投影片的標題 */
function getSectionHeaders(slides) {
  return slides
    .map((slide, index) => {
      if (slide.getLayout().getLayoutName() === 'SECTION_HEADER') {
        const title = getFirstTextboxText(slide);
        if (title) return { title, index, slideId: slide.getObjectId() };
      }
      return null;
    })
    .filter(Boolean);
}

/** 取第一個 textbox 的文字 */
function getFirstTextboxText(slide) {
  return slide.getShapes()
    .filter(s => s.getShapeType() === SlidesApp.ShapeType.TEXT_BOX)
    .map(s => s.getText().asString().trim())
    .find(t => t) || '';
}

/** Client-side 直接刪除舊的 TAB_LIST_* shapes/lines，改用 objectId 前綴判斷 */
function deleteOldTabsClient(slide) {
  slide.getShapes().forEach(shape => {
    const id = shape.getObjectId();
    if (id.startsWith('tab_') || id.startsWith('tab_bg_') || id.startsWith('page_num_')) { // Changed to match updated IDs
      shape.remove();
    }
  });
  // Also remove lines created with 'tab_line_' prefix
  slide.getLines().forEach(line => {
    const id = line.getObjectId();
    if (id.startsWith('tab_line_')) {
      line.remove();
    }
  });
}

/** 把一整列標籤的 batchUpdate requests 推到 requests 陣列 */
function appendTabListToSlide({ slideId, requests, sections, currentSection, config }) {
  // Use xStart directly from config, as it's already calculated
  let xPos = config.xStart;

  addBackgroundTabBar(slideId, requests, config);

  sections.forEach((sec, idx) => {
    const isActive = idx === currentSection;
    // Use the width and height pre-calculated in tabData
    appendTab({
      slideId,
      requests,
      title: sec.title,
      targetSlideId: sec.slideId,
      xPos,
      width: sec.width,
      height: config.height, // Use uniform height
      config,
      textColor: isActive ? '#FFFFFF' : config.inactiveTextColor,
      fillColor: isActive ? config.mainColor : config.bgColor
    });
    xPos += sec.width + config.spacing;
  });

  // Add the bottom line. It should be below all tabs.
  addBottomLine(slideId, requests, config);
}

/** 批次建立頁碼文字 */
function appendPageNumberToSlide({ slideId, requests, currentPage, totalPages, config }) {
  const pageNumId = `page_num_${slideId}_${newGuid()}`;
  requests.push(
    // 建立分頁文字框
    {
      createShape: {
        objectId: pageNumId,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: slideId,
          size: { height: { magnitude: 30, unit: 'PT' }, width: { magnitude: 50, unit: 'PT' } },
          transform: { translateX: 665, translateY: 370, scaleX: 1, scaleY: 1, unit: 'PT' }
        }
      }
    },
    // 插入分頁文字
    { insertText: { objectId: pageNumId, text: `${currentPage} / ${totalPages}` } },
    // 文字樣式 & 對齊
    {
      updateTextStyle: {
        objectId: pageNumId,
        textRange: { type: 'ALL' },
        style: {
          bold: true,
          fontFamily: config.mainFont,
          fontSize: { magnitude: 12, unit: 'PT' },
          foregroundColor: { opaqueColor: { rgbColor: hexToRgb(config.inactiveTextColor) } }
        },
        fields: 'bold,fontFamily,fontSize,foregroundColor'
      }
    },
    {
      updateParagraphStyle: {
        objectId: pageNumId,
        textRange: { type: 'ALL' },
        style: { alignment: 'CENTER' },
        fields: 'alignment'
      }
    }
  );
}

function appendTab({ slideId, requests, title, targetSlideId, xPos, width, height, config, textColor, fillColor }) { // Added height parameter
  const tabId = `tab_${slideId}_${newGuid()}`;
  requests.push(
    {
      createShape: {
        objectId: tabId,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: slideId,
          size: { height: { magnitude: height, unit: 'PT' }, width: { magnitude: width, unit: 'PT' } }, // Use calculated height
          transform: { translateX: xPos, translateY: config.y, scaleX: 1, scaleY: 1, unit: 'PT' }
        }
      }
    },
    { insertText: { objectId: tabId, text: title } },
    { updateShapeProperties: { objectId: tabId, shapeProperties: { shapeBackgroundFill: solidFill(fillColor), contentAlignment: 'MIDDLE' }, fields: 'shapeBackgroundFill.solidFill.color,contentAlignment' } },
    { updateTextStyle: { objectId: tabId, textRange: { type: 'ALL' }, style: { bold: true, fontFamily: config.mainFont, fontSize: { magnitude: config.fontSize, unit: 'PT' }, foregroundColor: { opaqueColor: { rgbColor: hexToRgb(textColor) } }, underline: false, link: { pageObjectId: targetSlideId } }, fields: 'bold,fontFamily,fontSize,foregroundColor,underline,link' } },
    { updateParagraphStyle: { objectId: tabId, textRange: { type: 'ALL' }, style: { alignment: 'CENTER' }, fields: 'alignment' } }
  );
}

function addBackgroundTabBar(slideId, requests, config) {
  const bgId = `tab_bg_${slideId}_${newGuid()}`;
  requests.push(
    {
      createShape: {
        objectId: bgId,
        shapeType: 'RECTANGLE',
        elementProperties: {
          pageObjectId: slideId,
          size: { height: { magnitude: config.height, unit: 'PT' }, width: { magnitude: config.totalWidth, unit: 'PT' } }, // Use uniform height
          transform: { translateX: 0, translateY: config.y, scaleX: 1, scaleY: 1, unit: 'PT' }
        }
      }
    },
    { updateShapeProperties: { objectId: bgId, shapeProperties: { shapeBackgroundFill: solidFill(config.bgColor), outline: { weight: { magnitude: 0.1, unit: 'PT' }, outlineFill: { solidFill: { color: { rgbColor: hexToRgb(config.bgColor) } } } } }, fields: 'shapeBackgroundFill.solidFill.color,outline.weight,outline.outlineFill.solidFill.color' } }
  );
}

function addBottomLine(slideId, requests, config) {
  const lineId = `tab_line_${slideId}_${newGuid()}`;
  requests.push(
    {
      createLine: {
        objectId: lineId,
        lineCategory: 'STRAIGHT',
        elementProperties: {
          pageObjectId: slideId,
          size: { height: { magnitude: 0, unit: 'PT' }, width: { magnitude: config.totalWidth, unit: 'PT' } },
          transform: { translateX: 0, translateY: config.y + config.height, scaleX: 1, scaleY: 1, unit: 'PT' } // Adjust Y based on new uniform height
        }
      }
    },
    { updateLineProperties: { objectId: lineId, lineProperties: { lineFill: solidFill(config.mainColor) }, fields: 'lineFill.solidFill.color' } }
  );
}

function solidFill(hex) {
  return { solidFill: { color: { rgbColor: hexToRgb(hex) } } };
}

function hexToRgb(hex) {
  const m = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return m ? { red: parseInt(m[1], 16) / 255, green: parseInt(m[2], 16) / 255, blue: parseInt(m[3], 16) / 255 } : { red: 0, green: 0, blue: 0 };
}

function newGuid() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 8);
}