const handouts = {
  CSST106_Lab_Experiment_01_Images: {
    title: "Lab 01 · Introduction to Computer Vision",
  },
  CSST106_Lab_Experiment_02_Edge_Detection: {
    title: "Lab 02 · Edge Detection",
  },
  CSST106_Lab_Experiment_03_Histogram: {
    title: "Lab 03 · Histogram Equalization",
  },
  CSST106_Lab_Experiment_04_SURF_Feature_Detection: {
    title: "Lab 04 · SURF Feature Detection",
  },
  CSST106_Lab_Experiment_05_PCA_in_CV: {
    title: "Lab 05 · PCA Facial Recognition",
  },
  CSST106_Lab_Experiment_06_YOLOV8_Object_Detection: {
    title: "Lab 06 · YOLO Object Detection",
  },
  CSST106_Lab_Experiment_07_Segmentation: {
    title: "Lab 07 · Image Segmentation",
  },
};

async function repairDanglingImageReferences(documentData) {
  const zip = await JSZip.loadAsync(documentData);
  const relationshipNamespace =
    "http://schemas.openxmlformats.org/package/2006/relationships";
  let repaired = 0;
  let placeholderData;

  for (const path of Object.keys(zip.files)) {
    if (!path.startsWith("word/") || !path.endsWith(".rels")) {
      continue;
    }

    const relationshipsFile = zip.file(path);
    if (!relationshipsFile) {
      continue;
    }

    const xmlText = await relationshipsFile.async("text");
    const relationshipsDocument = new DOMParser().parseFromString(
      xmlText,
      "application/xml",
    );
    if (relationshipsDocument.querySelector("parsererror")) {
      throw new Error(`Could not read DOCX relationships in ${path}.`);
    }

    const sourcePart = path
      .replace("/_rels/", "/")
      .replace(/\.rels$/, "");
    const relationships = [
      ...relationshipsDocument.getElementsByTagNameNS(
        relationshipNamespace,
        "Relationship",
      ),
    ];
    let updated = false;

    for (const relationship of relationships) {
      if (
        !relationship.getAttribute("Type")?.endsWith("/image") ||
        relationship.getAttribute("TargetMode") === "External"
      ) {
        continue;
      }

      const target = new URL(
        relationship.getAttribute("Target"),
        `https://docx.local/${sourcePart}`,
      ).pathname.replace(/^\/+/, "");
      const mediaFile = zip.file(target);
      if (!mediaFile || mediaFile.dir) {
        if (!placeholderData) {
          const canvas = document.createElement("canvas");
          canvas.width = 1;
          canvas.height = 1;
          placeholderData = canvas.toDataURL("image/png").split(",")[1];
        }

        const sourceDirectory = sourcePart.slice(
          0,
          sourcePart.lastIndexOf("/") + 1,
        );
        zip.file(
          `${sourceDirectory}media/viewer-missing-placeholder.png`,
          placeholderData,
          { base64: true },
        );
        relationship.setAttribute("Target", "media/viewer-missing-placeholder.png");
        repaired += 1;
        updated = true;
      }
    }

    if (updated) {
      zip.file(
        path,
        new XMLSerializer().serializeToString(relationshipsDocument),
      );
    }
  }

  // Repair dangling image relationships only in memory; the uploaded DOCX remains unchanged.
  const data =
    repaired > 0 ? await zip.generateAsync({ type: "arraybuffer" }) : documentData;
  return { data, repaired };
}

async function showHandout() {
  const params = new URLSearchParams(window.location.search);
  const filename = params.get("file") || "";
  const key = filename.replace(/\.docx$/i, "");
  const handout = handouts[key] || {
    title: key
      .replace(
        /^CSST106_Lab_Experiment_(\d+)_?/i,
        (_, number) => `Lab ${Number(number)} · `,
      )
      .replaceAll("_", " "),
  };
  const status = document.querySelector("#viewer-status");
  const title = document.querySelector("#viewer-title");
  const container = document.querySelector("#docx-container");
  const styleContainer = document.querySelector("#docx-styles");
  const originalLink = document.querySelector("#original-link");
  const zoomLevel = document.querySelector("#zoom-level");
  const warning = document.querySelector("#viewer-warning");
  let zoom = 1;
  let originalPageWidth = 0;

  function setZoom(value) {
    zoom = Math.max(0.4, Math.min(1.75, value));
    container.querySelectorAll(".docx").forEach((page) => {
      page.style.zoom = String(zoom);
    });
    zoomLevel.value = `${Math.round(zoom * 100)}%`;
  }

  function fitToWidth() {
    const wrapper = container.querySelector(".docx-wrapper");
    if (!wrapper || !originalPageWidth) {
      return;
    }

    const styles = window.getComputedStyle(wrapper);
    const horizontalPadding =
      Number.parseFloat(styles.paddingLeft) +
      Number.parseFloat(styles.paddingRight);
    const availableWidth = wrapper.clientWidth - horizontalPadding;
    setZoom(Math.min(1, availableWidth / originalPageWidth));
  }

  document.querySelector("#zoom-out").addEventListener("click", () => {
    setZoom(zoom - 0.1);
  });
  document.querySelector("#zoom-in").addEventListener("click", () => {
    setZoom(zoom + 0.1);
  });
  document.querySelector("#fit-width").addEventListener("click", fitToWidth);

  const isSafeFilename =
    /^[A-Za-z0-9][A-Za-z0-9 _.-]*\.docx$/i.test(filename) &&
    !filename.includes("..");
  if (!isSafeFilename) {
    status.textContent = "This handout could not be found.";
    status.classList.add("is-error");
    return;
  }

  title.textContent = handout.title;
  document.title = `${handout.title} | Vision Lab`;
  originalLink.href = `../${encodeURIComponent(filename)}`;
  originalLink.hidden = false;

  try {
    const response = await fetch(originalLink.href);
    if (!response.ok) {
      throw new Error(`Could not load the DOCX file (HTTP ${response.status}).`);
    }

    const documentData = await response.arrayBuffer();
    const preparedDocument = await repairDanglingImageReferences(documentData);
    await docx.renderAsync(preparedDocument.data, container, styleContainer, {
      className: "docx",
      inWrapper: true,
      breakPages: true,
      ignoreLastRenderedPageBreak: false,
      renderHeaders: true,
      renderFooters: true,
      renderFootnotes: true,
      renderEndnotes: true,
      useBase64URL: true,
    });
    const firstPage = container.querySelector(".docx");
    if (!firstPage) {
      throw new Error("The document renderer did not produce any pages.");
    }
    originalPageWidth = firstPage.getBoundingClientRect().width;
    if (preparedDocument.repaired > 0) {
      warning.textContent =
        `${preparedDocument.repaired} image reference(s) in this handout point to missing files. A blank placeholder preserves the layout; the original DOCX is unchanged.`;
      warning.hidden = false;
    }
    fitToWidth();
    status.hidden = true;
  } catch (error) {
    console.error("Unable to display the DOCX handout.", error);
    status.textContent =
      "The handout preview could not be loaded. You can open the original DOCX file instead.";
    status.classList.add("is-error");
  }
}

showHandout();
