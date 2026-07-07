(function () {
  "use strict";

  // Removed Firebase Config

  // HTMLの中にある様々な要素（ボタンや入力欄など）をJavaScriptで操作できるように取得します
  const csvInput = document.getElementById("csvInput");

  const excelOptions = document.getElementById("excelOptions");
  const sheetSelect = document.getElementById("sheetSelect");
  const tableSelect = document.getElementById("tableSelect");
  const tablePreview = document.getElementById("tablePreview");
  const chartType = document.getElementById("chartType");
  const legendPosition = document.getElementById("legendPosition");
  const reverseXAxis = document.getElementById("reverseXAxis");
  const xColumn = document.getElementById("xColumn");
  const yColumnContainer = document.getElementById("yColumnContainer");
  const chartTitle = document.getElementById("chartTitle");
  const xLabel = document.getElementById("xLabel");
  const yLabel = document.getElementById("yLabel");
  const yLabelRight = document.getElementById("yLabelRight");

  const savePngButton = document.getElementById("savePngButton");
  const savePdfButton = document.getElementById("savePdfButton");
  const saveXlsxButton = document.getElementById("saveXlsxButton");
  const saveXlsmButton = document.getElementById("saveXlsmButton");
  const message = document.getElementById("message");
  const canvas = document.getElementById("chartCanvas");
  const ctx = canvas.getContext("2d");
  const BASE_WIDTH = 1200;
  const BASE_HEIGHT = 760;

  // Removed AI and Auth elements

  const translations = {
    ja: {
      appTitle: "レポート用グラフ作成",
      appTitleMain: "レポート用グラフ作成",
      savePng: "PNG保存",
      savePdf: "PDF保存",
      saveXlsx: "XLSX保存",
      saveXlsm: "XLSM保存",
      loadCsvExcel: "CSV / Excel読み込み",
      chartType: "グラフ種類",
      lineChart: "折れ線グラフ",
      scatterPlot: "散布図",
      barChart: "棒グラフ",
      legendPosition: "凡例の位置",
      topRight: "右上",
      topLeft: "左上",
      bottomRight: "右下",
      bottomLeft: "左下",
      reverseXAxis: "X軸の左右を反転する",
      addSeries: "+ 系列を追加",
      title: "タイトル",
      chart: "グラフ",
      xLabel: "X軸ラベル",
      yLabelLeft: "左Y軸ラベル",
      yLabelRight: "右Y軸ラベル",
      createChart: "グラフ作成",
      selectFileError: "ファイルを選択してください。",
      readError: "ファイルの読み込みに失敗しました: ",
      noDataError: "描画するデータがありません。",
      xColumnRequired: "X軸の列を選択してください。",
      yColumnRequired: "少なくとも1つのY軸の列を選択してください。",
      loading: "読み込み中...",
      errorPrefix: "エラー: ",
      disclaimer: "※作成したグラフは必ずしも全ての論文誌の規定に合致するとは限りません。投稿先の規定をご確認ください。"
    },
    en: {
      appTitle: "Paper Graph Creator",
      appTitleMain: "Paper Graph Creator",
      savePng: "Save PNG",
      savePdf: "Save PDF",
      saveXlsx: "Save XLSX",
      saveXlsm: "Save XLSM",
      loadCsvExcel: "Load CSV / Excel",
      chartType: "Chart Type",
      lineChart: "Line Chart",
      scatterPlot: "Scatter Plot",
      barChart: "Bar Chart",
      legendPosition: "Legend Position",
      topRight: "Top Right",
      topLeft: "Top Left",
      bottomRight: "Bottom Right",
      bottomLeft: "Bottom Left",
      reverseXAxis: "Reverse X-Axis",
      addSeries: "+ Add Series",
      title: "Title",
      chart: "Chart",
      xLabel: "X-Axis Label",
      yLabelLeft: "Left Y-Axis Label",
      yLabelRight: "Right Y-Axis Label",
      createChart: "Create Chart",
      selectFileError: "Please select a file.",
      readError: "Failed to read file: ",
      noDataError: "No data to draw.",
      xColumnRequired: "Please select an X-axis column.",
      yColumnRequired: "Please select at least one Y-axis column.",
      loading: "Loading...",
      errorPrefix: "Error: ",
      disclaimer: "* The created graphs may not necessarily meet the requirements of all academic journals. Please check the submission guidelines."
    }
  };

  let currentLang = localStorage.getItem("appLang") || "ja";

  /**
   * アプリケーションの言語（日本語/英語）を切り替える関数です。
   * 画面上の特定の属性を持つ要素のテキストやプレースホルダーを指定された言語に翻訳します。
   */
  function applyLanguage(lang) {
    currentLang = lang;
    localStorage.setItem("appLang", lang);
    document.documentElement.lang = lang;

    document.querySelectorAll("[data-i18n]").forEach(el => {
      const key = el.getAttribute("data-i18n");
      if (translations[lang][key]) {
        el.textContent = translations[lang][key];
      }
    });

    document.querySelectorAll("[data-i18n-value]").forEach(el => {
      const key = el.getAttribute("data-i18n-value");
      if (translations[lang][key]) {
        el.value = translations[lang][key];
      }
    });

    document.querySelectorAll("[data-i18n-placeholder]").forEach(el => {
      const key = el.getAttribute("data-i18n-placeholder");
      if (translations[lang][key]) {
        el.placeholder = translations[lang][key];
      }
    });
  }

  const langSelect = document.getElementById("langSelect");
  if (langSelect) {
    langSelect.value = currentLang;
    langSelect.addEventListener("change", (e) => {
      applyLanguage(e.target.value);
    });
  }

  // 初期言語の適用
  applyLanguage(currentLang);

  // Removed Auth and Payment Status handling

  // Removed AI model event listeners

  let dataStores = [];
  let nextStoreId = 1;
  let currentFileName = "chart";
  let hasChart = false;
  let excelWorkbook = null;

  drawEmpty();

  csvInput.addEventListener("change", async () => {
    const file = csvInput.files && csvInput.files[0];
    if (!file) {
      return;
    }

    try {
      resetLoadedData();
      currentFileName = cleanFileName(file.name.replace(/\.[^.]+$/, "") || "chart");
      if (isExcelFile(file.name)) {
        excelWorkbook = await parseExcelWorkbook(await file.arrayBuffer());
        excelWorkbook.sheets.forEach((sheet) => {
          if (sheet.tables && sheet.tables.length > 0) {
            // 複数表がある場合、最後の1つは「シート全体」のデータなので除外する（重複を防ぐため）
            const tablesToLoad = sheet.tables.length > 1 ? sheet.tables.slice(0, -1) : sheet.tables;

            tablesToLoad.forEach((table, index) => {
              const tableName = tablesToLoad.length > 1 ? `${sheet.name} - 表${index + 1}` : sheet.name;
              applyTable(table.data, `${tableName} を読み込みました。`, `${currentFileName} - ${tableName}`);
            });
          }
        });
      } else {
        const text = await file.text();
        const parsed = parseCsv(text);
        applyTable(parsed, `CSVファイル (${currentFileName}) を読み込みました。`, currentFileName);
      }

      if (dataStores.length > 0 && seriesContainer.children.length === 0) {
        addSeriesCard(dataStores[0].id);
      }
    } catch (err) {
      setMessage(`エラー: ${err.message}`);
    }
  });
  //6/30 残り1810行
  const addSeriesBtn = document.getElementById("addSeriesBtn");
  const seriesContainer = document.getElementById("seriesContainer");
  if (addSeriesBtn) {
    addSeriesBtn.addEventListener("click", () => addSeriesCard());
  }

  [chartType, legendPosition, reverseXAxis, chartTitle, xLabel, yLabel, yLabelRight].forEach((control) => {
    if (control) {
      control.addEventListener("change", () => {
        if (dataStores.length) {
          drawChart();
        }
      });
    }
  });



  // Removed AI analyze button and markdown parsing

  savePngButton.addEventListener("click", () => {
    const link = document.createElement("a");
    link.download = `${currentFileName}.png`;
    link.href = canvas.toDataURL("image/png");
    link.click();
  });

  savePdfButton.addEventListener("click", () => {
    const dataUrl = canvas.toDataURL("image/jpeg", 0.95);
    const pdfBytes = createPdfWithJpeg(dataUrl, canvas.width, canvas.height);
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const link = document.createElement("a");
    link.download = `${currentFileName}.pdf`;
    link.href = URL.createObjectURL(blob);
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  });

  saveXlsxButton.addEventListener("click", () => {
    saveExcelWorkbook("xlsx");
  });

  saveXlsmButton.addEventListener("click", () => {
    saveExcelWorkbook("xlsm");
  });

  /**
   * グラフ作成ボタンや系列追加ボタンなどのコントロール部品の有効・無効を切り替えます。
   * データが読み込まれているかどうかに応じて操作可否を制御します。
   */
  function enableControls(enabled) {
    [chartType, legendPosition, reverseXAxis, chartTitle, xLabel, yLabel, yLabelRight, addSeriesBtn].forEach((control) => {
      if (control) control.disabled = !enabled;
    });
  }

  /**
   * 読み込まれたExcel/CSVデータ、および設定されたグラフ系列などの状態を初期化します。
   * 新しいファイルを読み込む前や、データをクリアする際に実行されます。
   */
  function resetLoadedData() {
    excelWorkbook = null;
    dataStores = [];
    nextStoreId = 1;
    hasChart = false;
    enableControls(false);
    setSaveButtons(false);
    seriesContainer.innerHTML = "";

  }
  //7/2残り1760行
  /**
   * 解析された二次元配列データを正規化し、ヘッダー行とデータ行に分類します。
   * その後、正規化されたデータをグラフ描画用データストアに保存するための処理を呼び出します。
   */
  function applyTable(parsedRaw, summaryText, sourceName) {
    const parsed = normalizeTable(parsedRaw);
    if (parsed.length < 2) {
      throw new Error("見出し行とデータ行が必要です。");
    }

    const headers = parsed[0].map((value, index) => String(value).replace(/^\uFEFF/, "").trim() || `列${index + 1}`);
    const rows = parsed.slice(1).filter((row) => row.some((value) => String(value).trim() !== ""));

    handleDataLoaded(summaryText, headers, rows, sourceName);
  }

  /**
   * 読み込みに成功したテーブルデータをデータストアに追加し、
   * 画面上のコントロールを有効化して、一時メッセージをクリアします。
   */
  function handleDataLoaded(summaryText, headers, rows, sourceName) {
    if (!rows || rows.length === 0) return;

    const storeId = "store_" + (nextStoreId++);
    dataStores.push({ id: storeId, name: sourceName || currentFileName, headers, rows });

    enableControls(true);
    hasChart = false;
    setSaveButtons(false);

    setMessage("");

    // Excelファイルで複数シートがある場合、すべての読み込みが終わってから最初のカードを作成するため、
    // ここでの自動生成は行いません（csvInputのイベントリスナー側で生成します）。
  }
  //7/6残り1720行
  /**
   * グラフの描画系列（X軸・Y軸・表示軸・近似直線など）を設定するための
   * UIカードを画面に追加し、各種セレクトボックスの変更イベントをハンドリングします。
   */
  function addSeriesCard(defaultStoreId = null) {
    if (dataStores.length === 0) return;
    const storeId = defaultStoreId || dataStores[0].id;
    const ds = dataStores.find(d => d.id === storeId);
    if (!ds) return;

    const card = document.createElement("div");
    card.className = "series-card";

    const headerDiv = document.createElement("div");
    headerDiv.className = "series-header";
    headerDiv.textContent = "データ系列";
    const removeBtn = document.createElement("button");
    removeBtn.className = "series-remove-btn";
    removeBtn.textContent = "×";
    removeBtn.onclick = () => { card.remove(); drawChart(); };
    headerDiv.appendChild(removeBtn);

    //7/6残り1700

    const row1 = document.createElement("div");
    row1.className = "series-row";
    const srcSelect = document.createElement("select");
    srcSelect.className = "series-source";
    dataStores.forEach(d => {
      srcSelect.add(new Option(d.name, d.id));
    });
    srcSelect.value = storeId;
    const l1 = document.createElement("span"); l1.className = "series-label"; l1.textContent = "データ元:";
    row1.appendChild(l1);
    row1.appendChild(srcSelect);

    const row2 = document.createElement("div");
    row2.className = "series-row";
    const xSelect = document.createElement("select");
    xSelect.className = "series-x";
    const l2 = document.createElement("span"); l2.className = "series-label"; l2.textContent = "X列:";
    row2.appendChild(l2);
    row2.appendChild(xSelect);

    const row3 = document.createElement("div");
    row3.className = "series-row";
    const ySelect = document.createElement("select");
    ySelect.className = "series-y";
    const l3 = document.createElement("span"); l3.className = "series-label"; l3.textContent = "Y列:";
    row3.appendChild(l3);
    row3.appendChild(ySelect);

    const row4 = document.createElement("div");
    row4.className = "series-row series-sub-row";
    const axisSelect = document.createElement("select");
    axisSelect.className = "series-axis";
    axisSelect.add(new Option("左軸", "left"));
    axisSelect.add(new Option("右軸", "right"));
    const l4 = document.createElement("span"); l4.className = "series-label"; l4.textContent = "表示軸:";
    row4.appendChild(l4);
    row4.appendChild(axisSelect);

    const row6 = document.createElement("div");
    row6.className = "series-row";
    const trendSelect = document.createElement("select");
    trendSelect.className = "series-trend";
    trendSelect.add(new Option("近似(なし)", "none"));
    trendSelect.add(new Option("線形", "linear"));
    trendSelect.add(new Option("指数", "exponential"));
    trendSelect.add(new Option("対数", "logarithmic"));
    trendSelect.add(new Option("累乗", "power"));
    trendSelect.add(new Option("多項式", "polynomial"));
    const l6 = document.createElement("span"); l6.className = "series-label"; l6.textContent = "近似直線:";
    row6.appendChild(l6);
    row6.appendChild(trendSelect);

    card.appendChild(headerDiv);
    card.appendChild(row1);
    card.appendChild(row2);
    card.appendChild(row3);
    card.appendChild(row4);
    card.appendChild(row6);
    seriesContainer.appendChild(card);

    const updateCols = () => {
      const selectedDs = dataStores.find(d => d.id === srcSelect.value);
      xSelect.innerHTML = "";
      ySelect.innerHTML = "";
      if (selectedDs) {
        selectedDs.headers.forEach((h, i) => {
          xSelect.add(new Option(h, String(i)));
          ySelect.add(new Option(h, String(i)));
        });
        xSelect.value = "0";
        ySelect.value = String(Math.min(1, selectedDs.headers.length - 1));
      }
    };

    const syncLabels = (changedEl) => {
      const ds = dataStores.find(d => d.id === srcSelect.value);
      if (!ds) return;

      // xLabelは1つ目の系列のX列に合わせる
      if ((changedEl === xSelect || changedEl === srcSelect) && seriesContainer.children[0] === card) {
        xLabel.value = ds.headers[Number(xSelect.value)] || "X";
      }

      // yLabelは選択された軸に合わせて更新
      if (changedEl === ySelect || changedEl === axisSelect || changedEl === srcSelect) {
        // 全系列の中で、自分がその軸を使っている最初の系列かどうか判定
        const allCards = Array.from(seriesContainer.children);
        const myAxis = axisSelect.value;
        const firstCardOfThisAxis = allCards.find(c => c.querySelector(".series-axis").value === myAxis);

        if (firstCardOfThisAxis === card) {
          const label = ds.headers[Number(ySelect.value)] || "Y";
          if (myAxis === "left") yLabel.value = label;
          else yLabelRight.value = label;
        }
      }
    };

    updateCols();
    srcSelect.addEventListener("change", () => { updateCols(); syncLabels(srcSelect); drawChart(); });

    [xSelect, ySelect, axisSelect, trendSelect].forEach(el => {
      el.addEventListener("change", () => {
        syncLabels(el);
        drawChart();
      });
    });

    syncLabels(srcSelect);
    drawChart();
  }

  const COLORS = ["#2563eb", "#dc2626", "#16a34a", "#ca8a04", "#9333ea", "#0891b2", "#ea580c"];
  let currentChartSeries = [];

  /**
   * 設定された全てのデータ系列を読み込んで解析し、選択されたグラフの種類（散布図・棒・折れ線）
   * に応じて対応する描画関数を呼び出し、キャンバス上に描画します。
   */
  function drawChart() {
    try {
      if (dataStores.length === 0 || seriesContainer.children.length === 0) {
        throw new Error("データ系列を追加してください。");
      }

      const allSeries = [];
      let colorIndex = 0;

      Array.from(seriesContainer.children).forEach(card => {
        const srcId = card.querySelector(".series-source").value;
        const xIdx = Number(card.querySelector(".series-x").value);
        const yIdx = Number(card.querySelector(".series-y").value);
        const axis = card.querySelector(".series-axis").value;
        const trend = card.querySelector(".series-trend").value;

        const ds = dataStores.find(d => d.id === srcId);
        if (!ds) return;

        const legend = ds.headers[yIdx] || `系列${allSeries.length + 1}`;

        const points = [];
        ds.rows.forEach(row => {
          const xRaw = String(row[xIdx] ?? "").trim();
          const xVal = parseNumber(row[xIdx]);
          const yVal = parseNumber(row[yIdx]);
          if (Number.isFinite(yVal) && xRaw !== "") {
            points.push({ xRaw: xRaw, x: chartType.value === "scatter" ? xVal : xRaw, y: yVal });
          }
        });

        if (points.length > 0) {
          allSeries.push({
            name: legend,
            sourceId: srcId,
            xIndex: xIdx,
            yIndex: yIdx,
            color: COLORS[colorIndex % COLORS.length],
            points: points,
            axis: axis,
            trendType: trend
          });
          colorIndex++;
        }
      });

      if (allSeries.length === 0) {
        throw new Error("選択されたY列に数値データが見つかりません。");
      }

      const leftSeries = allSeries.filter(s => s.axis === "left");
      if (!yLabel.value && leftSeries.length > 0) {
        yLabel.value = leftSeries[0].name;
      }
      const rightSeries = allSeries.filter(s => s.axis === "right");
      if (!yLabelRight.value && rightSeries.length > 0) {
        yLabelRight.value = rightSeries[0].name;
      }

      if (chartType.value === "scatter") {
        allSeries.forEach(s => {
          s.points = s.points.filter(p => Number.isFinite(p.x));
        });
        if (allSeries.every(s => s.points.length === 0)) {
          throw new Error("散布図ではX列も数値にしてください。");
        }
        renderScatter(allSeries.filter(s => s.points.length > 0));
      } else if (chartType.value === "bar") {
        renderBar(allSeries);
      } else {
        renderLine(allSeries);
      }

      currentChartSeries = allSeries;
      hasChart = true;
      setSaveButtons(true);
      setMessage("");
    } catch (error) {
      hasChart = false;
      setSaveButtons(false);
      setMessage(error.message);
    }
  }

  /**
   * PNG、PDF、XLSX、XLSM保存用ボタンの有効・無効を制御します。
   * グラフが正しく描画されている場合に保存を許可します。
   */
  function setSaveButtons(enabled) {
    [savePngButton, savePdfButton, saveXlsxButton, saveXlsmButton].forEach((button) => {
      button.disabled = !enabled;
    });
  }

  /**
   * グラフを描画する前段階としてキャンバスをクリアし、
   * 論文用のグラフ外枠および下部中央にタイトル（図番号）を描画します。
   */
  function renderFrame(titleText) {
    clearCanvas();
    // 余白をさらに広げる (left: 120->160, right: 120->160, bottom: 140->160)
    const plot = { left: 160, top: 40, right: BASE_WIDTH - 160, bottom: BASE_HEIGHT - 160 };
    ctx.fillStyle = "#000000";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.font = "700 24px 'Times New Roman', Times, serif";

    // 論文風に下部にタイトルを配置。プレフィックス「図: 」を付ける
    const textToDraw = titleText ? `図: ${titleText}` : "図: グラフ";
    ctx.fillText(textToDraw, BASE_WIDTH / 2, BASE_HEIGHT - 24);

    return plot;
  }

  /**
   * グラフ内の各系列に対応する色、シンボル、および近似直線の凡例を
   * 指定された表示位置（右上、左上、右下、左下）に合わせてキャンバス上に描画します。
   */
  function drawLegend(plot, seriesList) {
    const trendlineMap = {
      linear: "線形近似",
      exponential: "指数近似",
      logarithmic: "対数近似",
      power: "累乗近似",
      polynomial: "多項式近似"
    };

    const items = [];
    seriesList.forEach(s => {
      items.push({ name: s.name, type: chartType.value, color: s.color });
      if (s.trendType && s.trendType !== "none") {
        const trendName = trendlineMap[s.trendType] || "近似曲線";
        items.push({ name: `${s.name} (${trendName})`, type: "trendline", color: s.color });
      }
    });

    ctx.save();
    ctx.font = "14px 'Times New Roman', Times, serif";
    let maxWidth = 0;
    items.forEach(item => {
      const w = ctx.measureText(item.name).width;
      if (w > maxWidth) maxWidth = w;
    });

    const boxWidth = maxWidth + 60;
    const boxHeight = items.length * 24 + 16;

    let boxX, boxY;
    switch (legendPosition.value) {
      case "top-left":
        boxX = plot.left + 12;
        boxY = plot.top + 12;
        break;
      case "bottom-right":
        boxX = plot.right - boxWidth - 12;
        boxY = plot.bottom - boxHeight - 12;
        break;
      case "bottom-left":
        boxX = plot.left + 12;
        boxY = plot.bottom - boxHeight - 12;
        break;
      case "top-right":
      default:
        boxX = plot.right - boxWidth - 12;
        boxY = plot.top + 12;
        break;
    }

    ctx.fillStyle = "#ffffff";
    ctx.fillRect(boxX, boxY, boxWidth, boxHeight);
    ctx.strokeStyle = "#000000";
    ctx.lineWidth = 1;
    ctx.setLineDash([]);
    ctx.strokeRect(boxX, boxY, boxWidth, boxHeight);

    ctx.fillStyle = "#000000";
    ctx.textAlign = "left";
    ctx.textBaseline = "middle";

    items.forEach((item, i) => {
      const itemY = boxY + 20 + i * 24;
      const textX = boxX + 44;
      ctx.fillStyle = "#000000";
      ctx.fillText(item.name, textX, itemY);

      const symX = boxX + 22;
      if (item.type === "scatter") {
        drawPoint(symX, itemY, item.color, 4);
      } else if (item.type === "line") {
        ctx.beginPath();
        ctx.moveTo(symX - 12, itemY);
        ctx.lineTo(symX + 12, itemY);
        ctx.strokeStyle = item.color;
        ctx.lineWidth = 2;
        ctx.stroke();
        drawPoint(symX, itemY, "#ffffff", 3, item.color);
      } else if (item.type === "bar") {
        ctx.fillStyle = item.color;
        ctx.fillRect(symX - 8, itemY - 6, 16, 12);
        ctx.strokeStyle = "#000000";
        ctx.lineWidth = 1.5;
        ctx.strokeRect(symX - 8, itemY - 6, 16, 12);
      } else if (item.type === "trendline") {
        ctx.beginPath();
        ctx.moveTo(symX - 12, itemY);
        ctx.lineTo(symX + 12, itemY);
        ctx.strokeStyle = item.color;
        ctx.lineWidth = 2;
        ctx.setLineDash([6, 4]);
        ctx.stroke();
        ctx.setLineDash([]);
      }
    });
    ctx.restore();
  }

  /**
   * 指定された軸（左Y軸または右Y軸）に対応する系列データから、
   * 軸の目盛り範囲（最小値・最大値・ステップ）を計算して返します。
   */
  function getScaleForAxis(seriesList, axis) {
    const axisSeries = seriesList.filter(s => s.axis === axis);
    if (axisSeries.length === 0) return null;
    const ys = [];
    axisSeries.forEach(s => s.points.forEach(p => ys.push(p.y)));
    return makeScale(ys);
  }

  /**
   * キャンバス上に折れ線グラフを描画します。
   * 各データ系列のプロットを結ぶ線を引き、マーカーを描画し、近似直線が必要であれば重ねて描画します。
   */
  function renderLine(seriesList) {
    const plot = renderFrame(chartTitle.value);

    const yScaleLeft = getScaleForAxis(seriesList, "left");
    const yScaleRight = getScaleForAxis(seriesList, "right");
    const yScalePrimary = yScaleLeft || yScaleRight;

    // Assume all series share the same X raw labels (use the first one's X raw)
    const xLabels = seriesList[0].points.map(point => point.xRaw);
    drawAxes(plot, yScaleLeft, yScaleRight, xLabels);

    const xAt = (index, length) => {
      if (length <= 1) {
        return (plot.left + plot.right) / 2;
      }
      const ratio = index / (length - 1);
      const effectiveRatio = (reverseXAxis && reverseXAxis.checked) ? (1 - ratio) : ratio;
      return plot.left + effectiveRatio * (plot.right - plot.left);
    };
    const yAt = (value, axis) => {
      const scale = axis === "right" && yScaleRight ? yScaleRight : yScalePrimary;
      return plot.bottom - ((value - scale.min) / (scale.max - scale.min)) * (plot.bottom - plot.top);
    };

    let anyTrend = false;

    seriesList.forEach(series => {
      ctx.beginPath();
      series.points.forEach((point, index) => {
        const x = xAt(index, series.points.length);
        const y = yAt(point.y, series.axis);
        if (index === 0) {
          ctx.moveTo(x, y);
        } else {
          ctx.lineTo(x, y);
        }
      });
      ctx.strokeStyle = series.color;
      ctx.lineWidth = 2;
      ctx.stroke();

      series.points.forEach((point, index) => drawPoint(xAt(index, series.points.length), yAt(point.y, series.axis), "#ffffff", 5, series.color));

      const scaleToUse = series.axis === "right" && yScaleRight ? yScaleRight : yScalePrimary;
      const hasTrend = calculateAndDrawTrendline(series.points, plot, null, scaleToUse, series.color, series.trendType);
      if (hasTrend) anyTrend = true;
    });

    drawLegend(plot, seriesList);
  }

  /**
   * キャンバス上に散布図を描画します。
   * X列およびY列の数値データをプロット位置に変換し、点マーカーおよび近似直線を描画します。
   */
  function renderScatter(seriesList) {
    const plot = renderFrame(chartTitle.value);

    const allXs = [];
    seriesList.forEach(s => s.points.forEach(p => allXs.push(p.x)));
    const xScale = makeScale(allXs);

    const yScaleLeft = getScaleForAxis(seriesList, "left");
    const yScaleRight = getScaleForAxis(seriesList, "right");
    const yScalePrimary = yScaleLeft || yScaleRight;

    drawAxes(plot, yScaleLeft, yScaleRight, makeTickLabels(xScale), xScale);

    let anyTrend = false;

    seriesList.forEach(series => {
      series.points.forEach((point) => {
        const ratio = (point.x - xScale.min) / (xScale.max - xScale.min);
        const effectiveRatio = (reverseXAxis && reverseXAxis.checked) ? (1 - ratio) : ratio;
        const x = plot.left + effectiveRatio * (plot.right - plot.left);
        const scale = series.axis === "right" && yScaleRight ? yScaleRight : yScalePrimary;
        const y = plot.bottom - ((point.y - scale.min) / (scale.max - scale.min)) * (plot.bottom - plot.top);
        drawPoint(x, y, series.color, 5, series.color);
      });

      const scaleToUse = series.axis === "right" && yScaleRight ? yScaleRight : yScalePrimary;
      const hasTrend = calculateAndDrawTrendline(series.points, plot, xScale, scaleToUse, series.color, series.trendType);
      if (hasTrend) anyTrend = true;
    });

    drawLegend(plot, seriesList);
  }

  /**
   * キャンバス上に棒グラフを描画します。
   * 複数の系列がある場合はカテゴリごとにグループ化（クラスター化）して並べた棒を描画します。
   */
  function renderBar(seriesList) {
    const plot = renderFrame(chartTitle.value);

    const getBarScale = (axis) => {
      const scale = getScaleForAxis(seriesList, axis);
      if (scale) scale.min = Math.min(0, scale.min);
      return scale;
    };
    const yScaleLeft = getBarScale("left");
    const yScaleRight = getBarScale("right");
    const yScalePrimary = yScaleLeft || yScaleRight;

    const xLabels = seriesList[0].points.map(point => point.xRaw);
    drawAxes(plot, yScaleLeft, yScaleRight, xLabels);

    const gap = 8;
    const numPoints = seriesList[0].points.length;
    const numSeries = seriesList.length;
    const clusterWidth = Math.max(8, (plot.right - plot.left) / numPoints - gap);
    const barWidth = clusterWidth / numSeries;

    seriesList.forEach((series, sIndex) => {
      const scale = series.axis === "right" && yScaleRight ? yScaleRight : yScalePrimary;
      const zeroY = plot.bottom - ((0 - scale.min) / (scale.max - scale.min)) * (plot.bottom - plot.top);

      series.points.forEach((point, index) => {
        const ratio = (index + 0.5) / numPoints;
        const effectiveRatio = (reverseXAxis && reverseXAxis.checked) ? (1 - ratio) : ratio;
        const clusterCenter = plot.left + effectiveRatio * (plot.right - plot.left);
        const barCenter = clusterCenter - (clusterWidth / 2) + (sIndex * barWidth) + (barWidth / 2);

        const y = plot.bottom - ((point.y - scale.min) / (scale.max - scale.min)) * (plot.bottom - plot.top);
        const top = Math.min(y, zeroY);
        const height = Math.abs(zeroY - y);

        ctx.fillStyle = series.color;
        ctx.fillRect(barCenter - barWidth / 2, top, barWidth, Math.max(1, height));
        ctx.strokeStyle = "#000000";
        ctx.lineWidth = 1.5;
        ctx.strokeRect(barCenter - barWidth / 2, top, barWidth, Math.max(1, height));
      });
    });

    drawLegend(plot, seriesList);
  }

  /**
   * グラフのX軸、および左右のY軸の目盛り線、目盛り数値、軸ラベルを描画します。
   * 論文向けに目盛りを内側に向けて描画するなどの体裁を整えます。
   */
  function drawAxes(plot, yScaleLeft, yScaleRight, xLabels, xScale) {
    const drawYAxis = (scale, isRight) => {
      if (!scale) return;
      const yStep = scale.step || ((scale.max - scale.min) / 5);
      const yTicks = Math.max(1, Math.round((scale.max - scale.min) / yStep));

      ctx.font = "16px 'Times New Roman', Times, serif";
      ctx.fillStyle = "#000000";
      ctx.textAlign = isRight ? "left" : "right";
      ctx.textBaseline = "middle";

      for (let index = 0; index <= yTicks; index += 1) {
        const value = scale.min + index * yStep;
        const y = plot.bottom - (index / yTicks) * (plot.bottom - plot.top);

        ctx.strokeStyle = "#000000";
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        const tickX = isRight ? plot.right : plot.left;
        const tickDir = isRight ? -6 : 6;
        ctx.moveTo(tickX, y);
        ctx.lineTo(tickX + tickDir, y);
        ctx.stroke();

        const textX = isRight ? plot.right + 10 : plot.left - 10;
        ctx.fillText(formatNumber(value), textX, y);
      }

      const labelText = isRight ? yLabelRight.value : yLabel.value;
      if (labelText) {
        ctx.save();
        ctx.fillStyle = "#000000";
        ctx.font = "700 18px 'Times New Roman', Times, serif";
        // ラベルと値の距離をさらに広げる (80->110)
        ctx.translate(isRight ? plot.right + 110 : plot.left - 110, (plot.top + plot.bottom) / 2);
        ctx.rotate(isRight ? Math.PI / 2 : -Math.PI / 2);
        ctx.textAlign = "center";
        ctx.textBaseline = "middle";
        ctx.fillText(labelText, 0, 0);
        ctx.restore();
      }
    };

    ctx.strokeStyle = "#000000";
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.rect(plot.left, plot.top, plot.right - plot.left, plot.bottom - plot.top);
    ctx.stroke();

    drawYAxis(yScaleLeft, false);
    drawYAxis(yScaleRight, true);

    const maxLabels = Math.min(8, xLabels.length);
    const labelStep = Math.max(1, Math.ceil(xLabels.length / maxLabels));
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    for (let index = 0; index < xLabels.length; index += labelStep) {
      const ratio = xScale
        ? (xLabels[index] - xScale.min) / (xScale.max - xScale.min)
        : index / Math.max(1, xLabels.length - 1);
      const effectiveRatio = (reverseXAxis && reverseXAxis.checked) ? (1 - ratio) : ratio;
      const x = plot.left + effectiveRatio * (plot.right - plot.left);

      ctx.strokeStyle = "#000000";
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(x, plot.bottom);
      ctx.lineTo(x, plot.bottom - 6);
      ctx.moveTo(x, plot.top);
      ctx.lineTo(x, plot.top + 6);
      ctx.stroke();

      drawRotatedLabel(String(xLabels[index]), x, plot.bottom + 10);
    }

    ctx.save();
    ctx.fillStyle = "#000000";
    ctx.font = "700 18px 'Times New Roman', Times, serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    // X軸ラベルを少し上に移動 (40->75) して下の図タイトルと被らないようにする
    ctx.fillText(xLabel.value || "X", (plot.left + plot.right) / 2, canvas.height - 75);
    ctx.restore();
  }

  /**
   * X軸のラベル文字列が重ならないように、指定された座標で
   * 斜め（反時計回りに36度）に回転させてキャンバス上に描画します。
   */
  function drawRotatedLabel(text, x, y) {
    let value = text;
    if (typeof text === 'string' && text.toLowerCase().includes('e') && !isNaN(Number(text))) {
      value = formatNumber(Number(text));
    } else if (text && text.length > 14) {
      value = `${text.slice(0, 13)}...`;
    }
    
    ctx.save();
    ctx.translate(x, y);
    ctx.rotate(-Math.PI / 5);
    ctx.fillText(value, 0, 0);
    ctx.restore();
  }

  /**
   * キャンバス上の指定された座標（x, y）に、指定された色と半径で
   * プロット点（円形）を描画します。
   */
  function drawPoint(x, y, fillColor, radius, strokeColor) {
    ctx.beginPath();
    ctx.arc(x, y, radius, 0, Math.PI * 2);
    ctx.fillStyle = fillColor;
    ctx.fill();
    ctx.strokeStyle = strokeColor || "#000000";
    ctx.lineWidth = 1.5;
    ctx.stroke();
  }

  /**
   * キャンバスを真っ白にクリアした上で、データが読み込まれていない時の
   * プレースホルダー文字列（案内メッセージ）を中央に表示します。
   */
  function drawEmpty(text) {
    clearCanvas();
    ctx.fillStyle = "#000000";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.font = "700 24px 'Times New Roman', Times, serif";
    ctx.fillText(text || "CSVまたはExcelを読み込むとグラフを作成できます。", BASE_WIDTH / 2, BASE_HEIGHT / 2);
  }

  /**
   * キャンバス全体を消去し、背景を白色で塗りつぶします。
   * グラフの再描画を行う直前の初期化処理として使用されます。
   */
  function clearCanvas() {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  }

  /**
   * CSV形式のテキストデータを解析し、セル値を格納した二次元配列に変換します。
   * ダブルクォーテーションで囲まれた値やカンマ、改行のパースを行います。
   */
  function parseCsv(text) {
    const rowsOut = [];
    let row = [];
    let value = "";
    let inQuotes = false;

    for (let index = 0; index < text.length; index += 1) {
      const char = text[index];
      const next = text[index + 1];

      if (char === '"' && inQuotes && next === '"') {
        value += '"';
        index += 1;
      } else if (char === '"') {
        inQuotes = !inQuotes;
      } else if (char === "," && !inQuotes) {
        row.push(value);
        value = "";
      } else if ((char === "\n" || char === "\r") && !inQuotes) {
        if (char === "\r" && next === "\n") {
          index += 1;
        }
        row.push(value);
        rowsOut.push(row);
        row = [];
        value = "";
      } else {
        value += char;
      }
    }

    if (value !== "" || row.length) {
      row.push(value);
      rowsOut.push(row);
    }

    return rowsOut;
  }

  /**
   * ExcelやCSVから読み込んだ生データを整形し、結合されたヘッダー行を1行にまとめ、
   * 不要な空行や集計行（平均、標準偏差など）を除去して返します。
   */
  function normalizeTable(table) {
    const source = table
      .map((row) => row.map((value) => String(value ?? "").trim()))
      .filter((row) => row.some((value) => value !== ""));
    if (source.length < 2) {
      return source;
    }

    const dataStart = source.findIndex((row) => countNumericCells(row) >= 2 && !isSummaryRow(row));
    if (dataStart <= 0) {
      return source;
    }

    const headerRows = source
      .slice(0, dataStart)
      .filter((row) => countFilledCells(row) > 1);
    const effectiveHeaderRows = headerRows.length ? headerRows : [source[dataStart - 1]];
    const width = Math.max(...source.map((row) => row.length));
    const mergedHeaders = [];

    for (let columnIndex = 0; columnIndex < width; columnIndex += 1) {
      const parts = [];
      effectiveHeaderRows.forEach((row) => {
        const value = row[columnIndex] || "";
        if (value && parts[parts.length - 1] !== value) {
          parts.push(value);
        }
      });
      mergedHeaders[columnIndex] = parts.join(" / ") || `列${columnIndex + 1}`;
    }

    const bodyRows = source
      .slice(dataStart)
      .filter((row) => !isSummaryRow(row))
      .map((row) => {
        const next = Array.from({ length: width }, (_, index) => row[index] ?? "");
        return trimTrailingEmpty(next);
      })
      .filter((row) => row.some((value) => value !== ""));

    return [trimTrailingEmpty(mergedHeaders), ...bodyRows];
  }

  /**
   * 1つのシートデータ内から、数値データが多く含まれる「表」の候補領域を自動検出します。
   * 検出された各表候補と、シート全体のデータをそれぞれ二次元配列としてリスト化して返します。
   */
  function detectTableCandidates(rawTable) {
    const source = rawTable
      .map((row) => row.map((value) => String(value ?? "").trim()))
      .filter((row) => row.some((value) => value !== ""));
    const candidates = [];
    let index = 0;

    while (index < source.length) {
      while (index < source.length && (countNumericCells(source[index]) < 2 || isSummaryRow(source[index]))) {
        index += 1;
      }
      if (index >= source.length) {
        break;
      }

      const dataStart = index;
      while (index < source.length && (countNumericCells(source[index]) >= 2 || isSummaryRow(source[index]))) {
        index += 1;
      }
      const dataEnd = index;
      const headerStart = findHeaderStart(source, dataStart);
      const normalized = normalizeTable(source.slice(headerStart, dataEnd));
      addTableCandidate(candidates, normalized, headerStart, dataStart, dataEnd);
    }

    const wholeSheet = normalizeTable(source);
    addTableCandidate(candidates, wholeSheet, 0, 0, source.length);
    return candidates;
  }

  /**
   * 自動検出したデータ開始行の直前にある、表のヘッダー（列名）となる
   * 行の開始位置を上にさかのぼって探索してインデックスを返します。
   */
  function findHeaderStart(rows, dataStart) {
    let headerStart = dataStart;
    for (let index = dataStart - 1; index >= 0 && dataStart - index <= 6; index -= 1) {
      if (countNumericCells(rows[index]) >= 2 && !isSummaryRow(rows[index])) {
        break;
      }
      if (countFilledCells(rows[index]) >= 1) {
        headerStart = index;
      }
    }
    return headerStart;
  }

  /**
   * 抽出されたテーブルデータを重複チェックした上で、行数・列数などのメタデータと共に
   * 検出候補リスト（candidates）に追加します。
   */
  function addTableCandidate(candidates, table, headerStart, dataStart, dataEnd) {
    if (table.length < 2 || table[0].length < 2 || !table.slice(1).some((row) => countNumericCells(row) >= 1)) {
      return;
    }

    const signature = JSON.stringify(table);
    if (candidates.some((candidate) => candidate.signature === signature)) {
      return;
    }

    candidates.push({
      data: table,
      signature,
      label: `表${candidates.length + 1}: ${table.length - 1}行 x ${table[0].length}列 (Excel ${dataStart + 1}-${dataEnd}行付近)`
    });
  }

  /**
   * 指定された行の中で、空文字やスペース以外で値が入力されている
   * セルの個数を数えて返します。
   */
  function countFilledCells(row) {
    return row.filter((value) => String(value ?? "").trim() !== "").length;
  }

  /**
   * 指定された行の中で、数値として解析・認識できる有効な値が格納されている
   * セルの個数を数えて返します。
   */
  function countNumericCells(row) {
    return row.filter((value) => Number.isFinite(parseNumber(value))).length;
  }

  /**
   * 指定された行の先頭セルが「平均」や「標準偏差」といった
   * 集計やサマリーを示す行であるかどうかを判定します。
   */
  function isSummaryRow(row) {
    const label = String(row[0] ?? "").trim();
    return label === "平均" || label === "標準偏差";
  }

  /**
   * ファイル名からExcel形式（.xlsx / .xlsm）であるかどうかを判定します。
   * 古い .xls 形式の場合は未対応のエラーをスローします。
   */
  function isExcelFile(fileName) {
    const lowerName = fileName.toLowerCase();
    if (lowerName.endsWith(".xls")) {
      throw new Error("古い.xls形式は未対応です。.xlsx形式で保存し直して読み込んでください。");
    }
    return lowerName.endsWith(".xlsx") || lowerName.endsWith(".xlsm");
  }

  /**
   * Excelファイル（.xlsx等）のArrayBufferを解凍・パースし、
   * 各シート名と、その中に含まれる「表」のデータを解析してワークブック構造として返します。
   */
  async function parseExcelWorkbook(arrayBuffer) {
    const entries = await unzipXlsx(arrayBuffer);
    const relsXml = textEntry(entries, "_rels/.rels");
    const workbookPath = normalizeZipPath(findOfficeDocumentPath(relsXml) || "xl/workbook.xml");
    const workbookXml = textEntry(entries, workbookPath);
    const workbookDir = workbookPath.includes("/") ? workbookPath.slice(0, workbookPath.lastIndexOf("/")) : "";
    const workbookRelsPath = `${workbookDir}/_rels/${workbookPath.split("/").pop()}.rels`;
    const workbookRels = parseRelationships(textEntry(entries, workbookRelsPath));
    const workbookDoc = parseXml(workbookXml);
    const sheetNodes = Array.from(workbookDoc.getElementsByTagName("sheet"));
    if (!sheetNodes.length) {
      throw new Error("Excelファイルにシートが見つかりません。");
    }

    const sharedStrings = parseSharedStrings(entries.get("xl/sharedStrings.xml") || "");
    const sheets = sheetNodes.map((sheet, index) => {
      const relId = sheet.getAttribute("r:id") || sheet.getAttribute("id");
      const sheetTarget = workbookRels.get(relId);
      if (!sheetTarget) {
        return null;
      }
      const sheetPath = resolveZipPath(workbookDir, sheetTarget);
      const rawTable = parseSheetXml(textEntry(entries, sheetPath), sharedStrings);
      const tables = detectTableCandidates(rawTable);
      return {
        name: sheet.getAttribute("name") || `Sheet${index + 1}`,
        rawTable,
        tables
      };
    }).filter(Boolean);

    if (!sheets.some((sheet) => sheet.tables.length)) {
      throw new Error("Excel内にグラフ化できる表候補が見つかりません。");
    }

    return {
      fileName: currentFileName,
      sheets: sheets.filter((sheet) => sheet.tables.length)
    };
  }

  /**
   * Excelファイルの実体であるZIP圧縮データをバイナリレベルでパースし、
   * XMLファイル群を展開して、ファイル名と中身のテキストマップとして返します。
   */
  async function unzipXlsx(arrayBuffer) {
    const data = new Uint8Array(arrayBuffer);
    const view = new DataView(arrayBuffer);
    const eocdOffset = findEndOfCentralDirectory(data);
    const totalEntries = view.getUint16(eocdOffset + 10, true);
    const centralDirOffset = view.getUint32(eocdOffset + 16, true);
    const decoder = new TextDecoder("utf-8");
    const entries = new Map();
    let offset = centralDirOffset;

    for (let index = 0; index < totalEntries; index += 1) {
      if (view.getUint32(offset, true) !== 0x02014b50) {
        throw new Error("Excelファイルの構造を読み取れません。");
      }

      const method = view.getUint16(offset + 10, true);
      const compressedSize = view.getUint32(offset + 20, true);
      const fileNameLength = view.getUint16(offset + 28, true);
      const extraLength = view.getUint16(offset + 30, true);
      const commentLength = view.getUint16(offset + 32, true);
      const localHeaderOffset = view.getUint32(offset + 42, true);
      const nameBytes = data.slice(offset + 46, offset + 46 + fileNameLength);
      const name = decoder.decode(nameBytes);
      const normalizedName = normalizeZipPath(name);
      offset += 46 + fileNameLength + extraLength + commentLength;
      if (!normalizedName.endsWith(".xml") && !normalizedName.endsWith(".rels")) {
        continue;
      }

      const localNameLength = view.getUint16(localHeaderOffset + 26, true);
      const localExtraLength = view.getUint16(localHeaderOffset + 28, true);
      const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
      const compressed = data.slice(dataStart, dataStart + compressedSize);
      if (method !== 0 && method !== 8) {
        throw new Error("このExcelファイルの圧縮形式には対応していません。");
      }
      const bytes = method === 0 ? compressed : await inflateRaw(compressed);
      entries.set(normalizedName, decoder.decode(bytes));
    }

    return entries;
  }

  /**
   * ZIPファイル構造における「EOCD（セントラルディレクトリ終端レコード）」のオフセット位置を
   * バイナリデータの後方からスキャンして探し出します。
   */
  function findEndOfCentralDirectory(data) {
    for (let index = data.length - 22; index >= Math.max(0, data.length - 66000); index -= 1) {
      if (data[index] === 0x50 && data[index + 1] === 0x4b && data[index + 2] === 0x05 && data[index + 3] === 0x06) {
        return index;
      }
    }
    throw new Error("Excelファイルの終端情報が見つかりません。");
  }

  /**
   * DecompressionStreamを用いて、ZIP内のDeflate圧縮されたバイナリデータを
   * 生のバイナリ（解凍された状態）へ展開します。
   */
  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === "undefined") {
      throw new Error("このブラウザではExcelファイルの展開に対応していません。最新版のChromeまたはEdgeで開いてください。");
    }
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }

  /**
   * 解凍されたZIPエントリのマップから、指定されたパスのファイルを抽出し、
   * テキストデータとして返します。見つからない場合はエラーを発生させます。
   */
  function textEntry(entries, path) {
    const value = entries.get(normalizeZipPath(path));
    if (value === undefined) {
      throw new Error(`Excelファイル内の${path}を読み込めません。`);
    }
    return value;
  }

  /**
   * 文字列形式のXMLドキュメントをDOMParserでパースし、
   * JavaScriptで走査可能なXMLのDOMオブジェクトに変換して返します。
   */
  function parseXml(text) {
    const doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) {
      throw new Error("Excelファイル内のXMLを読み込めません。");
    }
    return doc;
  }

  /**
   * Excelの全体関係定義ファイル（.rels）から、主要なワークブックの実体となる
   * XMLファイルの格納パスを検索して取得します。
   */
  function findOfficeDocumentPath(relsXml) {
    const doc = parseXml(relsXml);
    const relationships = doc.getElementsByTagName("Relationship");
    for (const relationship of relationships) {
      if ((relationship.getAttribute("Type") || "").endsWith("/officeDocument")) {
        return relationship.getAttribute("Target");
      }
    }
    return "";
  }

  /**
   * XML要素の「Relationship」ノード群を解析し、
   * 各ID（Id）と関連ターゲットファイル（Target）のマッピングを保持するMapオブジェクトを構築します。
   */
  function parseRelationships(relsXml) {
    const doc = parseXml(relsXml);
    const result = new Map();
    for (const relationship of doc.getElementsByTagName("Relationship")) {
      result.set(relationship.getAttribute("Id"), relationship.getAttribute("Target"));
    }
    return result;
  }

  /**
   * Excel内で共有されている文字列テーブル（sharedStrings.xml）を解析し、
   * インデックス順に並べた文字列配列を構築して返します。
   */
  function parseSharedStrings(xmlText) {
    if (!xmlText) {
      return [];
    }
    const doc = parseXml(xmlText);
    return Array.from(doc.getElementsByTagName("si")).map((item) => {
      const textNodes = item.getElementsByTagName("t");
      return Array.from(textNodes)
        .filter((node) => {
          const parentName = node.parentNode ? node.parentNode.nodeName.toLowerCase() : "";
          return parentName !== "rph";
        })
        .map((node) => node.textContent || "")
        .join("");
    });
  }

  /**
   * Excelの各シートを表すXMLファイルを解析し、セル位置やマージセル情報を考慮して、
   * 行×列の二次元配列形式データへと変換します。
   */
  function parseSheetXml(xmlText, sharedStrings) {
    const doc = parseXml(xmlText);
    const output = [];
    for (const row of doc.getElementsByTagName("row")) {
      const rowIndex = Math.max(0, Number(row.getAttribute("r") || output.length + 1) - 1);
      const values = [];
      for (const cell of row.getElementsByTagName("c")) {
        const ref = cell.getAttribute("r") || "";
        const columnIndex = ref ? columnNameToIndex(ref.replace(/[0-9]/g, "")) : values.length;
        values[columnIndex] = readCellValue(cell, sharedStrings);
      }
      output[rowIndex] = trimTrailingEmpty(values).map((value) => value ?? "");
    }
    applyMergedCells(output, doc);
    return output
      .map((row) => trimTrailingEmpty(row || []).map((value) => value ?? ""))
      .filter((row) => row.some((value) => String(value).trim() !== ""));
  }

  /**
   * Excel内で結合されたセルの情報を読み取り、
   * 結合範囲内の空欄セルに対して左上セルの値を複製・補完する処理を行います。
   */
  function applyMergedCells(rows, doc) {
    for (const mergeCell of doc.getElementsByTagName("mergeCell")) {
      const ref = mergeCell.getAttribute("ref") || "";
      const [startRef, endRef] = ref.split(":");
      if (!startRef || !endRef) {
        continue;
      }

      const start = parseCellRef(startRef);
      const end = parseCellRef(endRef);
      const value = rows[start.row]?.[start.column] ?? "";
      if (value === "") {
        continue;
      }

      for (let rowIndex = start.row; rowIndex <= end.row; rowIndex += 1) {
        rows[rowIndex] = rows[rowIndex] || [];
        for (let columnIndex = start.column; columnIndex <= end.column; columnIndex += 1) {
          if (rows[rowIndex][columnIndex] === undefined || rows[rowIndex][columnIndex] === "") {
            rows[rowIndex][columnIndex] = value;
          }
        }
      }
    }
  }

  /**
   * Excelの「A1」や「B5」のようなセル座標参照文字列を解析し、
   * 0から始まる列インデックスと行インデックス of オブジェクトに変換します。
   */
  function parseCellRef(ref) {
    const column = ref.replace(/[0-9]/g, "");
    const row = ref.replace(/[A-Z]/gi, "");
    return {
      column: columnNameToIndex(column.toUpperCase()),
      row: Math.max(0, Number(row) - 1)
    };
  }

  /**
   * 各Excelセルの属性（型）を元に、共有文字列、真偽値、インライン文字列、
   * または数値を判定し、適切な文字列値を取り出して返します。
   */
  function readCellValue(cell, sharedStrings) {
    const type = cell.getAttribute("t");
    if (type === "inlineStr") {
      return Array.from(cell.getElementsByTagName("t")).map((node) => node.textContent || "").join("");
    }

    const valueNode = cell.getElementsByTagName("v")[0];
    const raw = valueNode ? valueNode.textContent || "" : "";
    if (type === "s") {
      return sharedStrings[Number(raw)] ?? "";
    }
    if (type === "b") {
      return raw === "1" ? "TRUE" : "FALSE";
    }
    return raw;
  }

  /**
   * Excelの列名記号（"A", "B", ..., "AA" など）を、
   * 0から始まる数値インデックス（A=0, B=1, ...）に変換して返します。
   */
  function columnNameToIndex(name) {
    let index = 0;
    for (const char of name) {
      index = index * 26 + char.charCodeAt(0) - 64;
    }
    return Math.max(0, index - 1);
  }

  /**
   * 配列の後方（末尾）に存在する、未定義値（undefined）や空文字列のセルを
   * 取り除いた状態の新しいスライス配列を返します。
   */
  function trimTrailingEmpty(values) {
    let end = values.length;
    while (end > 0 && (values[end - 1] === undefined || values[end - 1] === "")) {
      end -= 1;
    }
    return values.slice(0, end);
  }

  /**
   * ZIPアーカイブ内のファイルパスのバックスラッシュをスラッシュに置換し、
   * 先頭のスラッシュを除去して統一されたパス表記へと標準化します。
   */
  function normalizeZipPath(path) {
    return String(path || "").replace(/^\/+/, "").replace(/\\/g, "/");
  }

  /**
   * 基準となるディレクトリパス（baseDir）とターゲットの相対パスを結合し、
   * ".." などの参照を解決した正しいZIP内絶対パスを構築して返します。
   */
  function resolveZipPath(baseDir, target) {
    const normalizedTarget = normalizeZipPath(target);
    if (normalizedTarget.startsWith("xl/")) {
      return normalizedTarget;
    }
    const parts = `${baseDir}/${normalizedTarget}`.split("/");
    const resolved = [];
    for (const part of parts) {
      if (!part || part === ".") {
        continue;
      }
      if (part === "..") {
        resolved.pop();
      } else {
        resolved.push(part);
      }
    }
    return resolved.join("/");
  }

  /**
   * 文字列からカンマを取り除き、数値に変換します。
   * 空文字や数値に変換できない文字列の場合は NaN を返します。
   */
  function parseNumber(value) {
    if (value === null || value === undefined) {
      return NaN;
    }
    const cleaned = String(value).replace(/,/g, "").trim();
    if (!cleaned) {
      return NaN;
    }
    return Number(cleaned);
  }

  /**
   * 入力された数値群の最小値・最大値から、グラフ目盛りとして見栄えが良くなる
   * 最小値、最大値、および1目盛りあたりの間隔（ステップ）を決定します。
   */
  function makeScale(values) {
    let rawMin = Math.min(...values);
    let rawMax = Math.max(...values);
    if (!Number.isFinite(rawMin) || !Number.isFinite(rawMax)) {
      rawMin = 0;
      rawMax = 1;
    }

    if (rawMin === rawMax) {
      if (rawMin === 0) {
        rawMin = 0; rawMax = 10;
      } else {
        const pad = Math.abs(rawMin) * 0.1;
        rawMin -= pad;
        rawMax += pad;
      }
    } else {
      const range = rawMax - rawMin;
      let padMin = rawMin === 0 ? 0 : range * 0.05;
      let padMax = rawMax === 0 ? 0 : range * 0.05;
      if (rawMin > 0 && rawMin - padMin < 0) padMin = rawMin;
      if (rawMax < 0 && rawMax + padMax > 0) padMax = -rawMax;
      rawMin -= padMin;
      rawMax += padMax;
    }

    const targetTicks = 5;
    const roughStep = (rawMax - rawMin) / targetTicks;
    const mag = Math.pow(10, Math.floor(Math.log10(roughStep)));
    const normalizedStep = roughStep / mag;

    let niceStep;
    if (normalizedStep < 1.5) niceStep = 1 * mag;
    else if (normalizedStep < 2.5) niceStep = 2 * mag;
    else if (normalizedStep < 3.5) niceStep = 2.5 * mag;
    else if (normalizedStep < 7.5) niceStep = 5 * mag;
    else niceStep = 10 * mag;

    const min = Math.floor(rawMin / niceStep) * niceStep;
    const max = Math.ceil(rawMax / niceStep) * niceStep;

    return { min, max, step: niceStep };
  }

  /**
   * 算出されたスケール（min, max, step）に基づき、
   * グラフ軸に表示する目盛りのラベル値（単位表記付き）を配列として生成します。
   */
  function makeTickLabels(scale) {
    const step = scale.step || ((scale.max - scale.min) / 5);
    const ticks = Math.round((scale.max - scale.min) / step);
    return Array.from({ length: ticks + 1 }, (_, index) => {
      const value = scale.min + index * step;
      return formatNumber(value);
    });
  }

  /**
   * 数値に対して、M（メガ）、k（キロ）、m（ミリ）などの科学技術用の接頭辞を付与し、
   * 小数点以下を丸めて見やすい形式の文字列に変換します。
   */
  function formatNumber(value) {
    if (value === 0) return "0.00";
    
    // 常に有効数字3桁 (toExponential(2)) で指数表記のパーツを取得
    let str = value.toExponential(2);
    let parts = str.split('e');
    let coeff = parts[0];
    let exp = parseInt(parts[1], 10);
    
    // 指数が0の場合はそのまま係数のみ返す (例: 1.25)
    if (exp === 0) {
      return coeff;
    }
    
    // 指数部分を上付き文字に変換
    const superscripts = {
      '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
      '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹', '-': '⁻', '+': ''
    };
    let expStr = exp.toString().split('').map(c => superscripts[c] || c).join('');
    
    return coeff + ' × 10' + expStr;
  }

  /**
   * 画面上のエラーメッセージや状態表示エリアのテキストを指定された文字列で更新します。
   * メッセージがない場合はクリアします。
   */
  function setMessage(text) {
    message.textContent = text || "";
  }

  /**
   * 保存するファイル名に使えないWindowsの禁止文字（\, /, :, *, ?, ", <, >, |）を
   * アンダーバー (_) に置換して安全なファイル名を生成します。
   */
  function cleanFileName(name) {
    return name.replace(/[\\/:*?"<>|]/g, "_");
  }

  /**
   * 現在キャンバスに描画されているグラフデータ（系列データ）をメガテーブル形式に集約し、
   * XMLベースのExcelワークブックを生成して、ブラウザ経由でダウンロード保存します。
   */
  function saveExcelWorkbook(extension) {
    try {
      if (!hasChart || currentChartSeries.length === 0) {
        throw new Error("先にグラフを作成してください。");
      }

      const megaHeaders = [];
      const maxLength = Math.max(...currentChartSeries.map(s => s.points.length));
      const megaRows = Array.from({ length: maxLength }, () => []);

      currentChartSeries.forEach(s => {
        megaHeaders.push(s.name + "_X");
        megaHeaders.push(s.name);
      });

      for (let i = 0; i < maxLength; i++) {
        const row = [];
        currentChartSeries.forEach(s => {
          if (i < s.points.length) {
            row.push(s.points[i].x);
            row.push(s.points[i].y);
          } else {
            row.push("");
            row.push("");
          }
        });
        megaRows[i] = row;
      }

      const chartConfig = {
        series: currentChartSeries,
        chartType: chartType.value,
        title: chartTitle.value,
        xLabel: xLabel.value,
        yLabel: yLabel.value,
        rowCount: maxLength
      };

      const workbookBytes = createExcelWorkbook(extension, chartConfig, megaHeaders, megaRows);
      const type = extension === "xlsm"
        ? "application/vnd.ms-excel.sheet.macroEnabled.12"
        : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      const blob = new Blob([workbookBytes], { type });
      const link = document.createElement("a");
      link.download = `${currentFileName}.${extension}`;
      link.href = URL.createObjectURL(blob);
      link.click();
      setTimeout(() => URL.revokeObjectURL(link.href), 1000);
      setMessage("");
    } catch (error) {
      setMessage(error.message);
    }
  }

  /**
   * Excelワークブック（xlsx / xlsm）に必要なファイル構成一式（コンテンツタイプ定義、
   * リレーション、ワークシート、図面、グラフオブジェクト等）をZIPアーカイブとして組み立てます。
   */
  function createExcelWorkbook(extension, chartConfig, megaHeaders, megaRows) {
    const isMacroEnabled = extension === "xlsm";
    const sheetXml = createWorksheetXml(megaHeaders, megaRows);
    const workbookContentType = isMacroEnabled
      ? "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
      : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    return createZip({
      "[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="${workbookContentType}"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
  <Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`,
      "_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
      "xl/workbook.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="グラフ" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`,
      "xl/_rels/workbook.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
      "xl/worksheets/sheet1.xml": sheetXml,
      "xl/worksheets/_rels/sheet1.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>`,
      "xl/drawings/drawing1.xml": createDrawingXmlForChart(megaHeaders.length),
      "xl/drawings/_rels/drawing1.xml.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`,
      "xl/charts/chart1.xml": createChartXml(chartConfig)
    });
  }

  /**
   * Excelのワークシートデータ（sheet1.xml）のXML文字列を生成します。
   * ヘッダーと行データを各セル要素に変換し、グラフ配置用のdrawing要素との関連付けも定義します。
   */
  function createWorksheetXml(tableHeaders, tableRows) {
    const allRows = [tableHeaders, ...tableRows];
    const maxColumns = Math.max(1, ...allRows.map((row) => row.length));
    const dataRows = allRows.map((row, rowIndex) => {
      const cells = [];
      for (let columnIndex = 0; columnIndex < maxColumns; columnIndex += 1) {
        const value = row[columnIndex] ?? "";
        if (String(value).trim() === "") {
          continue;
        }
        const cellRef = `${columnIndexToName(columnIndex)}${rowIndex + 1}`;
        cells.push(createCellXml(cellRef, value, rowIndex === 0));
      }
      return `<row r="${rowIndex + 1}">${cells.join("")}</row>`;
    }).join("");
    const lastRef = `${columnIndexToName(Math.max(maxColumns - 1, 0))}${allRows.length}`;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:${lastRef}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>${Array.from({ length: maxColumns }, (_, index) => `<col min="${index + 1}" max="${index + 1}" width="16" customWidth="1"/>`).join("")}</cols>
  <sheetData>${dataRows}</sheetData>
  <drawing r:id="rId1"/>
</worksheet>`;
  }

  /**
   * Excelワークシート上の個別セルを表すXML断片を生成します。
   * 数値データは数値セル型として、文字列はインライン文字列型としてエスケープした上で格納します。
   */
  function createCellXml(cellRef, value, isHeader) {
    const text = String(value);
    const number = parseNumber(text);
    if (!isHeader && Number.isFinite(number)) {
      return `<c r="${cellRef}"><v>${number}</v></c>`;
    }
    return `<c r="${cellRef}" t="inlineStr"><is><t>${escapeXml(text)}</t></is></c>`;
  }

  /**
   * Excelのシート内にグラフオブジェクトを描画・配置するための
   * ドローイング定義ファイル（drawing1.xml）のXML文字列を生成します。
   */
  function createDrawingXmlForChart(columnCount) {
    const startColumn = Math.max(0, Math.min(columnCount + 1, 12));
    const widthEmu = 900 * 9525;
    const heightEmu = 570 * 9525;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
  <xdr:oneCellAnchor>
    <xdr:from>
      <xdr:col>${startColumn}</xdr:col>
      <xdr:colOff>0</xdr:colOff>
      <xdr:row>1</xdr:row>
      <xdr:rowOff>0</xdr:rowOff>
    </xdr:from>
    <xdr:ext cx="${widthEmu}" cy="${heightEmu}"/>
    <xdr:graphicFrame>
      <xdr:nvGraphicFramePr>
        <xdr:cNvPr id="2" name="グラフ"/>
        <xdr:cNvGraphicFramePr/>
      </xdr:nvGraphicFramePr>
      <xdr:xfrm>
        <a:off x="0" y="0"/>
        <a:ext cx="0" cy="0"/>
      </xdr:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
        </a:graphicData>
      </a:graphic>
    </xdr:graphicFrame>
    <xdr:clientData/>
  </xdr:oneCellAnchor>
</xdr:wsDr>`;
  }

  /**
   * Excelのグラフ実体ファイル（chart1.xml）のXML文字列を生成します。
   * グラフの種類や系列設定、系列に対応するデータ参照範囲、軸ラベルなどの定義を含みます。
   */
  function createChartXml(config) {
    const { series, chartType, title, xLabel, yLabel, rowCount } = config;
    const sheetName = "グラフ";

    const titleText = escapeXml(title || "グラフ");
    const xTitleText = escapeXml(xLabel || "X");
    const yTitleText = escapeXml(yLabel || "Y");

    let plotXml = "";

    const excelTrendMap = {
      linear: "linear",
      exponential: "exp",
      logarithmic: "log",
      power: "power",
      polynomial: "poly"
    };

    const seriesNodes = series.map((s, sIdx) => {
      // In the mega-table, each series takes 2 columns (X and Y).
      // So sIdx * 2 is X column, sIdx * 2 + 1 is Y column.
      const xCol = columnIndexToName(sIdx * 2);
      const yCol = columnIndexToName(sIdx * 2 + 1);

      const xRange = `='${sheetName}'!$${xCol}$2:$${xCol}$${s.points.length + 1}`;
      const yRange = `='${sheetName}'!$${yCol}$2:$${yCol}$${s.points.length + 1}`;

      const seriesName = escapeXml(s.name);
      const colorHex = COLORS[sIdx % COLORS.length].replace('#', '');

      const sTrendType = s.trendType;
      const trendNode = (sTrendType && sTrendType !== "none" && excelTrendMap[sTrendType]) ?
        `<c:trendline>
          <c:trendlineType val="${excelTrendMap[sTrendType]}"/>
          ${sTrendType === "polynomial" ? '<c:order val="2"/>' : ''}
          <c:dispEq val="0"/>
          <c:dispRSqr val="0"/>
        </c:trendline>` : "";

      if (chartType === "bar") {
        return `
        <c:ser>
          <c:idx val="${sIdx}"/>
          <c:order val="${sIdx}"/>
          <c:tx><c:v>${seriesName}</c:v></c:tx>
          <c:spPr><a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill></c:spPr>
          <c:cat><c:strRef><c:f>${xRange}</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>${yRange}</c:f></c:numRef></c:val>
          ${trendNode}
        </c:ser>`;
      } else {
        const spPr = chartType === "line"
          ? `<c:spPr><a:ln><a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill></a:ln></c:spPr>`
          : `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>`;
        const marker = `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill></a:ln></c:spPr></c:marker>`;
        return `
        <c:ser>
          <c:idx val="${sIdx}"/>
          <c:order val="${sIdx}"/>
          <c:tx><c:v>${seriesName}</c:v></c:tx>
          ${spPr}
          ${marker}
          <c:xVal><c:numRef><c:f>${xRange}</c:f></c:numRef></c:xVal>
          <c:yVal><c:numRef><c:f>${yRange}</c:f></c:numRef></c:yVal>
          ${trendNode}
        </c:ser>`;
      }
    }).join("");

    if (chartType === "bar") {
      plotXml = `
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:varyColors val="0"/>
        ${seriesNodes}
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:barChart>`;
    } else {
      const scatterStyle = chartType === "line" ? "lineMarker" : "marker";
      plotXml = `
      <c:scatterChart>
        <c:scatterStyle val="${scatterStyle}"/>
        <c:varyColors val="0"/>
        ${seriesNodes}
        <c:axId val="1"/>
        <c:axId val="2"/>
      </c:scatterChart>`;
    }

    const xAxisType = chartType === "bar" ? "c:catAx" : "c:valAx";
    const axesXml = `
      <${xAxisType}>
        <c:axId val="1"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="b"/>
        <c:majorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:majorGridlines>
        <c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/>
          <a:p><a:pPr><a:defRPr sz="1200" b="0"/></a:pPr><a:r><a:t>${xTitleText}</a:t></a:r></a:p>
          <a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:t>図: ${titleText}</a:t></a:r></a:p>
        </c:rich></c:tx>
        <c:layout/>
        <c:overlay val="0"/>
        </c:title>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="in"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:spPr><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>
        <c:txPr>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:pPr><a:defRPr sz="1100"/></a:pPr><a:endParaRPr/></a:p>
        </c:txPr>
        <c:crossAx val="2"/>
        <c:crosses val="autoZero"/>
      </${xAxisType}>
      <c:valAx>
        <c:axId val="2"/>
        <c:scaling><c:orientation val="minMax"/></c:scaling>
        <c:delete val="0"/>
        <c:axPos val="l"/>
        <c:majorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:majorGridlines>
        <c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1200" b="0"/></a:pPr><a:r><a:t>${yTitleText}</a:t></a:r></a:p></c:rich></c:tx>
        <c:layout/>
        <c:overlay val="0"/>
        </c:title>
        <c:numFmt formatCode="General" sourceLinked="1"/>
        <c:majorTickMark val="in"/>
        <c:minorTickMark val="none"/>
        <c:tickLblPos val="nextTo"/>
        <c:spPr><a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>
        <c:txPr>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p><a:pPr><a:defRPr sz="1100"/></a:pPr><a:endParaRPr/></a:p>
        </c:txPr>
        <c:crossAx val="1"/>
        <c:crosses val="autoZero"/>
      </c:valAx>
    `;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:plotArea>
      <c:spPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        <a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
      </c:spPr>
      ${plotXml}
      ${axesXml}
    </c:plotArea>
    <c:legend>
      <c:legendPos val="tr"/>
      <c:spPr>
        <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
        <a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
      </c:spPr>
    </c:legend>
  </c:chart>
</c:chartSpace>`;
  }

  /**
   * メモリ上で簡易的なZIPアーカイブを生成するためのバイナリ変換を行います。
   * 各ファイルのローカルヘッダー、セントラルディレクトリ、終端レコードを構築して結合します。
   */
  function createZip(files) {
    const encoder = new TextEncoder();
    const entries = Object.entries(files).map(([name, content]) => ({
      name,
      data: content instanceof Uint8Array ? content : encoder.encode(content)
    }));
    const parts = [];
    const centralParts = [];
    let offset = 0;
    const now = new Date();
    const dosTime = ((now.getHours() & 31) << 11) | ((now.getMinutes() & 63) << 5) | ((Math.floor(now.getSeconds() / 2)) & 31);
    const dosDate = (((now.getFullYear() - 1980) & 127) << 9) | (((now.getMonth() + 1) & 15) << 5) | (now.getDate() & 31);

    entries.forEach((entry) => {
      const nameBytes = encoder.encode(entry.name);
      const crc = crc32(entry.data);
      const localHeader = new Uint8Array(30 + nameBytes.length);
      const localView = new DataView(localHeader.buffer);
      localView.setUint32(0, 0x04034b50, true);
      localView.setUint16(4, 20, true);
      localView.setUint16(10, dosTime, true);
      localView.setUint16(12, dosDate, true);
      localView.setUint32(14, crc, true);
      localView.setUint32(18, entry.data.length, true);
      localView.setUint32(22, entry.data.length, true);
      localView.setUint16(26, nameBytes.length, true);
      localHeader.set(nameBytes, 30);
      parts.push(localHeader, entry.data);

      const centralHeader = new Uint8Array(46 + nameBytes.length);
      const centralView = new DataView(centralHeader.buffer);
      centralView.setUint32(0, 0x02014b50, true);
      centralView.setUint16(4, 20, true);
      centralView.setUint16(6, 20, true);
      centralView.setUint16(12, dosTime, true);
      centralView.setUint16(14, dosDate, true);
      centralView.setUint32(16, crc, true);
      centralView.setUint32(20, entry.data.length, true);
      centralView.setUint32(24, entry.data.length, true);
      centralView.setUint16(28, nameBytes.length, true);
      centralView.setUint32(42, offset, true);
      centralHeader.set(nameBytes, 46);
      centralParts.push(centralHeader);
      offset += localHeader.length + entry.data.length;
    });

    const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
    const endHeader = new Uint8Array(22);
    const endView = new DataView(endHeader.buffer);
    endView.setUint32(0, 0x06054b50, true);
    endView.setUint16(8, entries.length, true);
    endView.setUint16(10, entries.length, true);
    endView.setUint32(12, centralSize, true);
    endView.setUint32(16, offset, true);
    return concatBytes([...parts, ...centralParts, endHeader]);
  }

  /**
   * 与えられたバイト列（Uint8Array）に対して、
   * ZIPアーカイブのデータ整合性チェックに必要なCRC-32チェックサム値を算出して返します。
   */
  function crc32(bytes) {
    let crc = 0xffffffff;
    for (const byte of bytes) {
      crc = CRC_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  const CRC_TABLE = (() => {
    const table = [];
    for (let index = 0; index < 256; index += 1) {
      let value = index;
      for (let bit = 0; bit < 8; bit += 1) {
        value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
      }
      table[index] = value >>> 0;
    }
    return table;
  })();

  /**
   * 複数の Uint8Array（バイト配列）の断片を結合し、
   * 1つの連続した Uint8Array バイト配列として統合して返します。
   */
  function concatBytes(chunks) {
    const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const result = new Uint8Array(totalLength);
    let offset = 0;
    chunks.forEach((chunk) => {
      result.set(chunk, offset);
      offset += chunk.length;
    });
    return result;
  }

  /**
   * Base64でエンコードされたData URL（画像データなど）をデコードし、
   * 生のバイナリバイト配列（Uint8Array）に変換して返します。
   */
  function dataUrlToBytes(dataUrl) {
    const base64 = dataUrl.split(",")[1];
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return bytes;
  }

  /**
   * 0から始まる列インデックス（0, 1, 2...）を
   * Excelスタイルの列アルファベット記号（A, B, C...）に変換して返します。
   */
  function columnIndexToName(index) {
    let value = index + 1;
    let name = "";
    while (value > 0) {
      const remainder = (value - 1) % 26;
      name = String.fromCharCode(65 + remainder) + name;
      value = Math.floor((value - 1) / 26);
    }
    return name;
  }

  /**
   * XML/HTML内で特殊な意味を持つ5つの文字（&, <, >, ", '）を
   * 安全な実体参照形式（&amp;, &lt;, &gt;, &quot;, &#39;）にエスケープします。
   */
  function escapeXml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  /**
   * 文字列をHTMLエスケープします。内部的にはescapeXmlを流用し、
   * シングルクォートも含めてサニタイズ処理を行います。
   */
  function escapeHtml(value) {
    return escapeXml(value).replace(/'/g, "&#39;");
  }

  /**
   * グラフ画像を埋め込んだPDFドキュメントをバイナリレベルで組み立てて生成します。
   * JPEG画像をオブジェクトとして定義し、用紙サイズ内にスケーリングして配置します。
   */
  function createPdfWithJpeg(dataUrl, imageWidth, imageHeight) {
    const base64 = dataUrl.split(",")[1];
    const binary = atob(base64);
    const imageBytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      imageBytes[index] = binary.charCodeAt(index);
    }

    const pageWidth = 842;
    const pageHeight = 595;
    const margin = 36;
    const maxWidth = pageWidth - margin * 2;
    const maxHeight = pageHeight - margin * 2;
    const scale = Math.min(maxWidth / imageWidth, maxHeight / imageHeight);
    const drawWidth = imageWidth * scale;
    const drawHeight = imageHeight * scale;
    const drawX = (pageWidth - drawWidth) / 2;
    const drawY = (pageHeight - drawHeight) / 2;

    const imageCommand = `q\n${drawWidth.toFixed(2)} 0 0 ${drawHeight.toFixed(2)} ${drawX.toFixed(2)} ${drawY.toFixed(2)} cm\n/Im0 Do\nQ`;
    const objects = [
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${pageWidth} ${pageHeight}] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>`,
      `<< /Type /XObject /Subtype /Image /Width ${imageWidth} /Height ${imageHeight} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${imageBytes.length} >>\nstream\n${binary}\nendstream`,
      `<< /Length ${imageCommand.length} >>\nstream\n${imageCommand}\nendstream`
    ];

    let pdf = "%PDF-1.4\n";
    const offsets = [0];
    objects.forEach((object, index) => {
      offsets.push(pdf.length);
      pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
    });
    const xrefOffset = pdf.length;
    pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
    for (let index = 1; index <= objects.length; index += 1) {
      pdf += `${String(offsets[index]).padStart(10, "0")} 00000 n \n`;
    }
    pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;

    const bytes = new Uint8Array(pdf.length);
    for (let index = 0; index < pdf.length; index += 1) {
      bytes[index] = pdf.charCodeAt(index) & 0xff;
    }
    return bytes;
  }

  /**
   * 各系列のデータポイントから、選択された回帰方式（線形、対数、指数、累乗、多項式）の
   * 近似曲線を算出し、キャンバス上に破線で描画します。
   */
  function calculateAndDrawTrendline(points, plot, xScale, yScale, color, type) {
    if (!type || type === "none" || points.length < 2) return false;

    const numericPoints = points.map((p, i) => ({
      x: Number.isFinite(p.x) ? p.x : i,
      y: p.y
    }));

    let fn = null;
    let n = numericPoints.length;

    if (type === "linear") {
      let sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;
      numericPoints.forEach(p => {
        sumX += p.x; sumY += p.y; sumXY += p.x * p.y; sumXX += p.x * p.x;
      });
      const denominator = (n * sumXX - sumX * sumX);
      if (Math.abs(denominator) < 1e-10) return;
      const slope = (n * sumXY - sumX * sumY) / denominator;
      const intercept = (sumY - slope * sumX) / n;
      fn = (x) => slope * x + intercept;
    } else if (type === "exponential") {
      let sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;
      const validPoints = numericPoints.filter(p => p.y > 0);
      n = validPoints.length;
      if (n < 2) return;
      validPoints.forEach(p => {
        const ly = Math.log(p.y);
        sumX += p.x; sumY += ly; sumXY += p.x * ly; sumXX += p.x * p.x;
      });
      const denominator = (n * sumXX - sumX * sumX);
      if (Math.abs(denominator) < 1e-10) return;
      const b = (n * sumXY - sumX * sumY) / denominator;
      const a = Math.exp((sumY - b * sumX) / n);
      fn = (x) => a * Math.exp(b * x);
    } else if (type === "logarithmic") {
      let sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;
      const validPoints = numericPoints.filter(p => p.x > 0);
      n = validPoints.length;
      if (n < 2) return;
      validPoints.forEach(p => {
        const lx = Math.log(p.x);
        sumX += lx; sumY += p.y; sumXY += lx * p.y; sumXX += lx * lx;
      });
      const denominator = (n * sumXX - sumX * sumX);
      if (Math.abs(denominator) < 1e-10) return;
      const b = (n * sumXY - sumX * sumY) / denominator;
      const a = (sumY - b * sumX) / n;
      fn = (x) => a + b * Math.log(x);
    } else if (type === "power") {
      let sumX = 0, sumY = 0, sumXY = 0, sumXX = 0;
      const validPoints = numericPoints.filter(p => p.x > 0 && p.y > 0);
      n = validPoints.length;
      if (n < 2) return;
      validPoints.forEach(p => {
        const lx = Math.log(p.x);
        const ly = Math.log(p.y);
        sumX += lx; sumY += ly; sumXY += lx * ly; sumXX += lx * lx;
      });
      const denominator = (n * sumXX - sumX * sumX);
      if (Math.abs(denominator) < 1e-10) return;
      const b = (n * sumXY - sumX * sumY) / denominator;
      const a = Math.exp((sumY - b * sumX) / n);
      fn = (x) => a * Math.pow(x, b);
    } else if (type === "polynomial") {
      let sumX = 0, sumXX = 0, sumXXX = 0, sumXXXX = 0;
      let sumY = 0, sumXY = 0, sumXXY = 0;
      numericPoints.forEach(p => {
        const x2 = p.x * p.x;
        sumX += p.x; sumXX += x2; sumXXX += x2 * p.x; sumXXXX += x2 * x2;
        sumY += p.y; sumXY += p.x * p.y; sumXXY += x2 * p.y;
      });
      const det = n * (sumXX * sumXXXX - sumXXX * sumXXX) - sumX * (sumX * sumXXXX - sumXXX * sumXX) + sumXX * (sumX * sumXXX - sumXX * sumXX);
      if (Math.abs(det) < 1e-10) return;
      const detA = sumY * (sumXX * sumXXXX - sumXXX * sumXXX) - sumX * (sumXY * sumXXXX - sumXXY * sumXXX) + sumXX * (sumXY * sumXXX - sumXXY * sumXX);
      const detB = n * (sumXY * sumXXXX - sumXXX * sumXXY) - sumY * (sumX * sumXXXX - sumXXX * sumXX) + sumXX * (sumX * sumXXY - sumXX * sumXY);
      const detC = n * (sumXX * sumXXY - sumXY * sumXXX) - sumX * (sumX * sumXXY - sumXY * sumXX) + sumY * (sumX * sumXXX - sumXX * sumXX);
      const a = detA / det;
      const b = detB / det;
      const c = detC / det;
      fn = (x) => a + b * x + c * x * x;
    }

    if (!fn) return;

    ctx.save();
    ctx.beginPath();
    const segments = 200;
    const minX = Math.min(...numericPoints.map(p => p.x));
    const maxX = Math.max(...numericPoints.map(p => p.x));

    let started = false;
    for (let i = 0; i <= segments; i++) {
      const xVal = minX + (maxX - minX) * (i / segments);
      if (type === "logarithmic" || type === "power") {
        if (xVal <= 0) continue;
      }
      const yVal = fn(xVal);
      if (!Number.isFinite(yVal)) continue;

      let drawX, drawY;
      if (xScale) {
        const ratio = (xVal - xScale.min) / (xScale.max - xScale.min);
        const effectiveRatio = (reverseXAxis && reverseXAxis.checked) ? (1 - ratio) : ratio;
        drawX = plot.left + effectiveRatio * (plot.right - plot.left);
      } else {
        const ratio = (xVal - 0) / (points.length - 1);
        const effectiveRatio = (reverseXAxis && reverseXAxis.checked) ? (1 - ratio) : ratio;
        drawX = plot.left + effectiveRatio * (plot.right - plot.left);
      }
      drawY = plot.bottom - ((yVal - yScale.min) / (yScale.max - yScale.min)) * (plot.bottom - plot.top);

      // Prevent drawing way out of bounds
      if (drawY < plot.top - 1000 || drawY > plot.bottom + 1000) continue;

      if (!started) {
        ctx.moveTo(drawX, drawY);
        started = true;
      } else {
        ctx.lineTo(drawX, drawY);
      }
    }

    ctx.strokeStyle = color || "#000000";
    ctx.lineWidth = 2;
    ctx.setLineDash([8, 6]);
    ctx.stroke();
    ctx.restore();
    return true;
  }
})();
