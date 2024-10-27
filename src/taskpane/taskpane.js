let categoryData = {
  closing: [],
  postClosing: [],
  representation: [],
};
let allParagraphsData = [];
let isDataLoaded = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    const logStyleContentButton = document.getElementById("logStyleContentButton");
    const categorySelect = document.getElementById("categorySelect");

    logStyleContentButton.disabled = true;
    logStyleContentButton.onclick = getListInfoFromSelection;
    document.getElementById("clearContentButton").onclick = clearAllContent;

    categorySelect.onchange = handleCategoryChange;
    handleCategoryChange();

    loadAllParagraphsData();
  }
});

async function handleCategoryChange() {
  const categorySelect = document.getElementById("categorySelect");
  const selectedCategory = categorySelect.value;

  document.querySelectorAll(".category-content").forEach((section) => {
    section.classList.remove("active");
  });

  const contentId = `${selectedCategory}Content`;
  document.getElementById(contentId).classList.add("active");

  document.getElementById("logStyleContentButton").disabled = !isDataLoaded || !selectedCategory;

  // Add auto-copy functionality
  if (selectedCategory && categoryData[selectedCategory]) {
    const clipboardString = formatCategoryData(selectedCategory);
    await copyToClipboard(clipboardString);
  }
}

function normalizeText(text) {
  return text
    .trim()
    .replace(/\s+/g, " ")
    .replace(/[^\x20-\x7E]/g, "");
}

async function loadAllParagraphsData() {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      allParagraphsData = [];
      let parentNumbering = [];
      let lastNumbering = "";

      document.getElementById("logStyleContentButton").disabled = true;
      isDataLoaded = false;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text,isListItem");
        await context.sync();

        let text = normalizeText(paragraph.text);

        if (text.length <= 1) {
          continue;
        }

        if (paragraph.isListItem) {
          paragraph.listItem.load("level,listString");
          await context.sync();

          const level = paragraph.listItem.level;
          const listString = paragraph.listItem.listString || "";

          if (level <= parentNumbering.length) {
            parentNumbering = parentNumbering.slice(0, level);
          }

          parentNumbering[level] = listString;

          let fullNumbering = "";
          for (let j = 0; j <= level; j++) {
            if (parentNumbering[j]) {
              fullNumbering += `${parentNumbering[j]}.`;
            }
          }

          fullNumbering = fullNumbering.replace(/\.$/, "");
          lastNumbering = fullNumbering;

          allParagraphsData.push({
            key: fullNumbering,
            value: text,
            originalText: paragraph.text.trim(),
            isListItem: true,
            index: i,
            level: level,
            listString: listString,
            parentNumbers: [...parentNumbering],
          });
        } else {
          const key = lastNumbering ? `${lastNumbering} (text)` : `text_${i + 1}`;
          allParagraphsData.push({
            key: key,
            value: text,
            originalText: paragraph.text.trim(),
            isListItem: false,
            index: i,
            level: -1,
          });
        }
      }

      allParagraphsData = allParagraphsData.filter((item) => !item.key.endsWith(".text"));

      console.log("All paragraphs data loaded:", allParagraphsData);
      isDataLoaded = true;

      const categorySelect = document.getElementById("categorySelect");
      document.getElementById("logStyleContentButton").disabled = !categorySelect.value;
    });
  } catch (error) {
    console.error("An error occurred while loading all paragraphs data:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
    document.getElementById("logStyleContentButton").disabled = true;
    isDataLoaded = false;
  }
}

async function getListInfoFromSelection() {
  if (!isDataLoaded) {
    console.log("Data is still loading. Please wait.");
    return;
  }

  const selectedCategory = document.getElementById("categorySelect").value;
  if (!selectedCategory) {
    console.log("No category selected");
    return;
  }

  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const range = selection.getRange();
      range.load("text");
      await context.sync();

      const selectedRange = range.text;
      const paragraphs = selection.paragraphs;
      paragraphs.load("items");
      await context.sync();

      let newSelections = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const selectedParagraph = paragraphs.items[i];
        selectedParagraph.load("text,isListItem");
        await context.sync();

        if (selectedParagraph.isListItem) {
          selectedParagraph.listItem.load("level,listString");
          await context.sync();
        }

        const selectedText = selectedParagraph.text.trim();
        const normalizedSelectedText = normalizeText(selectedText);

        const matchingParagraphs = allParagraphsData.filter(
          (para) => para.value === normalizedSelectedText || para.originalText === selectedText
        );

        if (matchingParagraphs.length > 0) {
          let bestMatch = matchingParagraphs[0];

          if (matchingParagraphs.length > 1 && selectedParagraph.isListItem) {
            const selectedLevel = selectedParagraph.listItem.level;
            const selectedListString = selectedParagraph.listItem.listString;

            const exactMatch = matchingParagraphs.find(
              (para) => para.isListItem && para.level === selectedLevel && para.listString === selectedListString
            );

            if (exactMatch) {
              bestMatch = exactMatch;
            }
          }

          const isDuplicate = categoryData[selectedCategory].some(
            (item) => item.key === bestMatch.key && item.value === bestMatch.value
          );

          if (!isDuplicate) {
            newSelections.push({
              key: bestMatch.key,
              value: bestMatch.value,
            });
          }
        }
      }

      if (newSelections.length > 0) {
        categoryData[selectedCategory] = [...categoryData[selectedCategory], ...newSelections];

        categoryData[selectedCategory].sort((a, b) => {
          const aNumbers = a.key.split(".").map((num) => parseInt(num));
          const bNumbers = b.key.split(".").map((num) => parseInt(num));

          for (let i = 0; i < Math.max(aNumbers.length, bNumbers.length); i++) {
            if (isNaN(aNumbers[i])) return 1;
            if (isNaN(bNumbers[i])) return -1;
            if (aNumbers[i] !== bNumbers[i]) return aNumbers[i] - bNumbers[i];
          }
          return 0;
        });

        updateCategoryDisplay(selectedCategory);
        const clipboardString = formatCategoryData(selectedCategory);
        await copyToClipboard(clipboardString);

        console.log(`Updated ${selectedCategory} data:`, categoryData[selectedCategory]);
      }
    });
  } catch (error) {
    console.error("An error occurred while processing selection:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
  }
}

function formatCategoryData(category) {
  if (!categoryData[category] || !Array.isArray(categoryData[category])) {
    console.error("Invalid category data for:", category);
    return "{}";
  }

  const pairs = categoryData[category].map((pair) => `"${pair.key}": "${pair.value.replace(/"/g, '\\"')}"`).join(",\n");

  return `{\n${pairs}\n}`;
}

function updateCategoryDisplay(category) {
  const contentElement = document.querySelector(`#${category}Content .content-area`);
  if (!contentElement) {
    console.error("Content element not found for category:", category);
    return;
  }

  contentElement.innerHTML = "";

  if (categoryData[category]) {
    categoryData[category].forEach((pair) => {
      const keySpan = `<span class="key">${pair.key}</span>`;
      const valueSpan = `<span class="value">${pair.value}</span>`;
      const formattedPair = `<div class="pair">${keySpan}: ${valueSpan}</div>`;
      contentElement.innerHTML += formattedPair;
    });
  }
}

async function copyToClipboard(text) {
  if (!text) {
    console.error("No text provided to copy");
    showCopyMessage(false);
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    showCopyMessage(true);
  } catch (err) {
    console.log("Fallback: using execCommand for copy");
    const textArea = document.createElement("textarea");
    textArea.value = text;
    document.body.appendChild(textArea);

    try {
      textArea.select();
      const successful = document.execCommand("copy");
      showCopyMessage(successful);
    } catch (err) {
      console.error("Failed to copy text:", err);
      showCopyMessage(false);
    } finally {
      document.body.removeChild(textArea);
    }
  }
}

function showCopyMessage(successful) {
  const copyMessage = document.getElementById("copyMessage");
  if (!copyMessage) {
    console.error("Copy message element not found");
    return;
  }

  copyMessage.style.display = "block";
  copyMessage.textContent = successful ? "Content added and copied to clipboard!" : "Failed to copy content";
  copyMessage.style.color = successful ? "green" : "red";

  setTimeout(() => {
    copyMessage.style.display = "none";
  }, 3000);
}

function clearAllContent() {
  categoryData = {
    closing: [],
    postClosing: [],
    representation: [],
  };

  document.querySelectorAll(".content-area").forEach((element) => {
    element.innerHTML = "";
  });

  console.log("All content cleared");
}
