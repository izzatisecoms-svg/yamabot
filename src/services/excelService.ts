import XLSX from 'xlsx-js-style';
import { HotelRoomType, CompetitorHotel, CompetitorRoomMapping, PriceReportRow } from '../types/yama';
import { RapidApiService } from './rapidApiService';

export class ExcelService {
  /**
   * 1. Sheet: hotel_data
   * Columns: hotel_name | room_type | room_name | min_rate | max_rate
   */
  static buildHotelDataSheet(hotelName: string, hotelRooms: HotelRoomType[]): XLSX.WorkSheet {
    const rows: any[][] = [
      ['hotel_name', 'room_type', 'room_name', 'min_rate', 'max_rate'],
      ...hotelRooms.map(r => [
        hotelName,
        r.room_type,
        r.room_name || r.room_type,
        r.min_rate ?? 0,
        r.max_rate ?? 0
      ])
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 25 }, { wch: 18 }, { wch: 28 }, { wch: 14 }, { wch: 14 }];
    ws['!rows'] = [{ hpt: 24 }];

    // Apply header style
    const headerStyle = {
      fill: { fgColor: { rgb: "0B2545" } },
      font: { name: "Calibri", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
      alignment: { vertical: "center", horizontal: "left" },
      border: {
        top: { style: "thin", color: { rgb: "244F84" } },
        bottom: { style: "thin", color: { rgb: "244F84" } },
        left: { style: "thin", color: { rgb: "244F84" } },
        right: { style: "thin", color: { rgb: "244F84" } }
      }
    };
    for (let c = 0; c < 5; c++) {
      const ref = XLSX.utils.encode_cell({ r: 0, c });
      if (ws[ref]) ws[ref].s = headerStyle;
    }

    return ws;
  }

  /**
   * Enriches a competitor hotel record so url, address, city, currency_code, and zip are never blank
   */
  static enrichCompetitorHotel(h: CompetitorHotel): CompetitorHotel {
    const id = String(h.competitor_hotel_id || '').trim();
    const name = String(h.competitor_hotel_name || h.competitor_hotel_search || '').trim();
    const search = String(h.competitor_hotel_search || name).trim();
    const known = RapidApiService.findKnownHotel(id, name, search);

    const isIndo = (h.countrycode === 'id') || (h.currency_code === 'IDR') || (known?.countrycode === 'id');
    const currency = h.currency_code || (isIndo ? 'IDR' : 'MYR');
    const country = h.countrycode || known?.countrycode || (isIndo ? 'id' : 'my');

    const cleanSlug = (name || known?.competitor_hotel_name || id || 'hotel')
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, '-')
      .replace(/^-|-$/g, '');

    const url = (h.url && h.url.trim())
      ? h.url.trim()
      : (known?.url || `https://www.booking.com/hotel/${country}/${cleanSlug || id}.html`);

    const address = (h.address && h.address.trim())
      ? h.address.trim()
      : (known?.address || (isIndo ? 'Jalan Sudirman, Jakarta' : '76, Jalan Batu Unjur 1, Taman Bayu Perdana'));

    const city = (h.city && h.city.trim())
      ? h.city.trim()
      : (known?.city || (isIndo ? 'Jakarta' : 'Klang'));

    const zip = (h.zip && h.zip.trim())
      ? h.zip.trim()
      : (known?.zip || (isIndo ? '10110' : '41200'));

    return {
      competitor_hotel_search: search || known?.competitor_hotel_search || name,
      competitor_hotel_id: id || known?.competitor_hotel_id || '0',
      competitor_hotel_name: name || known?.competitor_hotel_name || `Hotel ${id}`,
      url,
      address,
      city,
      countrycode: country,
      currency_code: currency,
      zip,
      review_score: h.review_score ?? known?.review_score ?? 8.0
    };
  }

  /**
   * 2. Sheet: competitor_hotel
   * Exact format matching Image 1:
   * competitor_hotel_search | competitor_hotel_id | competitor_hotel_name | url | address | city | currency_code | zip
   */
  static buildCompetitorHotelSheet(competitorHotels: CompetitorHotel[]): XLSX.WorkSheet {
    const enrichedHotels = competitorHotels.map(h => this.enrichCompetitorHotel(h));
    const rows: any[][] = [
      ['competitor_hotel_search', 'competitor_hotel_id', 'competitor_hotel_name', 'url', 'address', 'city', 'currency_code', 'zip'],
      ...enrichedHotels.map(h => [
        h.competitor_hotel_search || h.competitor_hotel_name || '',
        h.competitor_hotel_id || '0',
        h.competitor_hotel_name || '',
        h.url || '',
        h.address || '',
        h.city || '',
        h.currency_code || (h.countrycode === 'id' ? 'IDR' : 'MYR'),
        h.zip || ''
      ])
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [
      { wch: 28 },
      { wch: 20 },
      { wch: 30 },
      { wch: 45 },
      { wch: 38 },
      { wch: 18 },
      { wch: 15 },
      { wch: 12 }
    ];
    ws['!rows'] = [{ hpt: 24 }];

    const headerStyle = {
      fill: { fgColor: { rgb: "0B2545" } },
      font: { name: "Calibri", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
      alignment: { vertical: "center", horizontal: "left" },
      border: {
        top: { style: "thin", color: { rgb: "244F84" } },
        bottom: { style: "thin", color: { rgb: "244F84" } },
        left: { style: "thin", color: { rgb: "244F84" } },
        right: { style: "thin", color: { rgb: "244F84" } }
      }
    };
    for (let c = 0; c < 8; c++) {
      const ref = XLSX.utils.encode_cell({ r: 0, c });
      if (ws[ref]) ws[ref].s = headerStyle;
    }

    return ws;
  }

  /**
   * 3. Sheet: competitor_room
   * Columns order:
   * hotel_name | competitor_hotel_name | competitor_hotel_id | competitor_room_id | competitor_room_name | currency_code | gross_amount | room_type | extracted_at
   */
  static buildCompetitorRoomSheet(hotelName: string, mappedRooms: CompetitorRoomMapping[]): XLSX.WorkSheet {
    const defaultTimestamp = new Date().toISOString().replace('T', ' ').substring(0, 19);
    const rows: any[][] = [
      ['hotel_name', 'competitor_hotel_name', 'competitor_hotel_id', 'competitor_room_id', 'competitor_room_name', 'currency_code', 'gross_amount', 'room_type', 'extracted_at'],
      ...mappedRooms.map(r => {
        let gross: any = '';
        if (typeof r.gross_amount === 'number' && !isNaN(r.gross_amount)) {
          gross = r.gross_amount;
        } else if (r.gross_amount !== undefined && r.gross_amount !== null && r.gross_amount !== '' && r.gross_amount !== 'N/A') {
          const parsed = Number(r.gross_amount);
          gross = isNaN(parsed) ? '' : parsed;
        }
        return [
          hotelName || '',
          r.competitor_hotel_name || '',
          r.competitor_hotel_id || '',
          r.competitor_room_id || '',
          r.competitor_room_name || '',
          r.currency_code || r.currency || 'MYR',
          gross,
          r.our_room_type || 'UNMAPPED',
          r.extracted_at || defaultTimestamp
        ];
      })
    ];
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [
      { wch: 25 },
      { wch: 28 },
      { wch: 20 },
      { wch: 20 },
      { wch: 35 },
      { wch: 15 },
      { wch: 15 },
      { wch: 18 },
      { wch: 22 }
    ];
    ws['!rows'] = [{ hpt: 24 }];

    const headerStyle = {
      fill: { fgColor: { rgb: "0B2545" } },
      font: { name: "Calibri", sz: 10, bold: true, color: { rgb: "FFFFFF" } },
      alignment: { vertical: "center", horizontal: "left" },
      border: {
        top: { style: "thin", color: { rgb: "244F84" } },
        bottom: { style: "thin", color: { rgb: "244F84" } },
        left: { style: "thin", color: { rgb: "244F84" } },
        right: { style: "thin", color: { rgb: "244F84" } }
      }
    };
    for (let c = 0; c < 9; c++) {
      const ref = XLSX.utils.encode_cell({ r: 0, c });
      if (ws[ref]) ws[ref].s = headerStyle;
    }

    return ws;
  }

  /**
   * 4. Sheet: competitor_price
   * Exact format matching Image 3 (Pivot matrix):
   * competitor_hotel_name | competitor_hotel_id | [room_type_1] | [room_type_2] | ...
   * Followed by summary rows:
   * Min-max rate
   * Average price
   * Lowest price
   */
  /**
   * 4. Sheet: competitor_price
   * Exact format matching image_93bc1049.png:
   * Title Banner: COMPETITOR PRICE REPORT | <hotelName>
   * Section 1: ROOM RATE SUMMARY | Quick pricing overview
   * Table 1: Metric | Description | [Room Types...]
   * Rows: Average, Lowest, Allowed range
   * Section 2: COMPETITOR PRICE COMPARISON | Room rates (RM)
   * Table 2: Competitor hotel | RapidAPI hotel ID | [Room Types...]
   * Rows: Each competitor hotel with formatted rates (e.g. RM 89.50) or N/A
   */
  static buildCompetitorPriceSheet(
    hotelRooms: HotelRoomType[],
    competitorHotels: CompetitorHotel[],
    mappedRooms: CompetitorRoomMapping[],
    priceReport?: PriceReportRow[],
    hotelName?: string
  ): XLSX.WorkSheet {
    // Unique room types from hotelRooms, or derived from priceReport / mappedRooms
    const roomTypes = hotelRooms.length > 0
      ? hotelRooms
      : (priceReport && priceReport.length > 0)
        ? priceReport.map(p => ({
            room_type: p.our_room_type,
            room_name: p.our_room_type,
            min_rate: p.our_min_rate,
            max_rate: p.our_max_rate
          }))
        : Array.from(new Set(mappedRooms.map(m => m.our_room_type).filter(Boolean)))
            .filter(t => !['UNMAPPED', 'IGNORE', 'NONE'].includes(t.toUpperCase()))
            .map(t => ({ room_type: t, room_name: t, min_rate: 0, max_rate: 0 }));

    // Detect currency (MYR -> RM, IDR -> Rp / IDR, etc.)
    const sampleCurrency = mappedRooms.find(m => m.currency_code || m.currency)?.currency_code ||
      mappedRooms.find(m => m.currency_code || m.currency)?.currency ||
      competitorHotels.find(h => h.currency_code)?.currency_code ||
      'MYR';

    const isIdr = String(sampleCurrency).toUpperCase() === 'IDR';
    const currSymbol = isIdr ? 'Rp ' : (String(sampleCurrency).toUpperCase() === 'USD' ? '$' : 'RM ');
    const currLabel = isIdr ? 'IDR' : (String(sampleCurrency).toUpperCase() === 'MYR' ? 'RM' : sampleCurrency);

    // Collect all competitor hotels
    const hotelsList: Array<{ id: string; name: string }> = [];
    const seenIds = new Set<string>();

    competitorHotels.forEach(h => {
      const id = String(h.competitor_hotel_id || '0').trim();
      const name = h.competitor_hotel_name || h.competitor_hotel_search || `Hotel ${id}`;
      const key = `${id}_${name}`;
      if (id && id !== '0' && !seenIds.has(key)) {
        seenIds.add(key);
        hotelsList.push({ id, name });
      }
    });

    mappedRooms.forEach(m => {
      const id = String(m.competitor_hotel_id || '0').trim();
      const name = m.competitor_hotel_name || `Hotel ${id}`;
      const key = `${id}_${name}`;
      if (id && id !== '0' && !seenIds.has(key)) {
        seenIds.add(key);
        hotelsList.push({ id, name });
      }
    });

    // Helper to format price
    const formatPrice = (val: number | null | undefined): string => {
      if (val === null || val === undefined || isNaN(val) || val <= 0) return 'N/A';
      return `${currSymbol}${val.toFixed(2)}`;
    };

    // Calculate rates per hotel and room type
    const pricesByRoomType: Record<string, number[]> = {};
    roomTypes.forEach(rt => {
      pricesByRoomType[rt.room_type] = [];
    });

    const hotelMatrixRows: any[][] = [];
    hotelsList.forEach(h => {
      const row: any[] = [h.name, h.id];
      roomTypes.forEach(rt => {
        const matchingRooms = mappedRooms.filter(m =>
          String(m.competitor_hotel_id || '').trim() === String(h.id).trim() &&
          String(m.our_room_type || '').trim().toUpperCase() === String(rt.room_type || '').trim().toUpperCase()
        );

        const validPrices: number[] = [];
        matchingRooms.forEach(m => {
          let val: number | null = null;
          if (typeof m.gross_amount === 'number' && !isNaN(m.gross_amount) && m.gross_amount > 0) {
            val = m.gross_amount;
          } else if (m.gross_amount && m.gross_amount !== 'N/A') {
            const parsed = Number(m.gross_amount);
            if (!isNaN(parsed) && parsed > 0) val = parsed;
          }
          if (val !== null) validPrices.push(val);
        });

        if (validPrices.length > 0) {
          const minPrice = Math.min(...validPrices);
          pricesByRoomType[rt.room_type].push(minPrice);
          row.push(formatPrice(minPrice));
        } else {
          row.push('N/A');
        }
      });
      hotelMatrixRows.push(row);
    });

    // Calculate Averages and Lowests
    const avgCells: string[] = roomTypes.map(rt => {
      const reportRow = priceReport?.find(p => p.our_room_type.toUpperCase() === rt.room_type.toUpperCase());
      if (reportRow && typeof reportRow.average_competitor_price === 'number') {
        return formatPrice(reportRow.average_competitor_price);
      }
      const collected = pricesByRoomType[rt.room_type] || [];
      if (collected.length > 0) {
        const sum = collected.reduce((a, b) => a + b, 0);
        return formatPrice(sum / collected.length);
      }
      return 'N/A';
    });

    const lowestCells: string[] = roomTypes.map(rt => {
      const reportRow = priceReport?.find(p => p.our_room_type.toUpperCase() === rt.room_type.toUpperCase());
      if (reportRow && typeof reportRow.lowest_competitor_price === 'number') {
        return formatPrice(reportRow.lowest_competitor_price);
      }
      const collected = pricesByRoomType[rt.room_type] || [];
      if (collected.length > 0) {
        return formatPrice(Math.min(...collected));
      }
      return 'N/A';
    });

    const allowedRangeCells: string[] = roomTypes.map(rt => {
      if (rt.min_rate !== undefined && rt.max_rate !== undefined && (rt.min_rate > 0 || rt.max_rate > 0)) {
        return `${rt.min_rate}–${rt.max_rate}`;
      }
      return 'N/A';
    });

    const totalCols = Math.max(roomTypes.length + 2, 8);

    // Build Rows exactly matching screenshot:
    // Row 1: COMPETITOR PRICE REPORT | <hotelName>
    const titleRow = [`COMPETITOR PRICE REPORT | ${hotelName || "Competitor Pricing Monitor"}`];
    // Row 2: ROOM RATE SUMMARY | Quick pricing overview
    const section1HeaderRow = ['ROOM RATE SUMMARY | Quick pricing overview'];
    // Row 3: Metric | Description | SS | DS | SD | DD | DT | FD | FP
    const table1Headers = ['Metric', 'Description', ...roomTypes.map(r => r.room_type)];
    // Row 4: Average | Valid competitor rates only | ...
    const avgRow = ['Average', 'Valid competitor rates only', ...avgCells];
    // Row 5: Lowest | Lowest valid rate | ...
    const lowestRow = ['Lowest', 'Lowest valid rate', ...lowestCells];
    // Row 6: Allowed range | Min–max configured | ...
    const rangeRow = ['Allowed range', 'Min–max configured', ...allowedRangeCells];
    // Row 7: Spacer
    const spacerRow: string[] = [];
    // Row 8: COMPETITOR PRICE COMPARISON | Room rates (RM)
    const section2HeaderRow = [`COMPETITOR PRICE COMPARISON | Room rates (${currLabel})`];
    // Row 9: Competitor hotel | Hotel ID | SS | DS | SD | DD | DT | FD | FP
    const table2Headers = ['Competitor hotel', 'Hotel ID', ...roomTypes.map(r => r.room_type)];
    // Row 10+: hotelMatrixRows

    const allRows = [
      titleRow,
      section1HeaderRow,
      table1Headers,
      avgRow,
      lowestRow,
      rangeRow,
      spacerRow,
      section2HeaderRow,
      table2Headers,
      ...hotelMatrixRows
    ];

    const ws = XLSX.utils.aoa_to_sheet(allRows);

    // Merges for the full-width banners
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: totalCols - 1 } }, // Row 1 title
      { s: { r: 1, c: 0 }, e: { r: 1, c: totalCols - 1 } }, // Row 2 summary header
      { s: { r: 7, c: 0 }, e: { r: 7, c: totalCols - 1 } }, // Row 8 comparison header
    ];

    // Column widths
    ws['!cols'] = [
      { wch: 28 }, // Metric / Competitor hotel
      { wch: 20 }, // Description / Hotel ID
      ...roomTypes.map(() => ({ wch: 16 }))
    ];

    // Row heights
    const rowHeights: Array<{ hpt: number }> = [
      { hpt: 32 }, // Row 1 Title
      { hpt: 24 }, // Row 2 Summary Header
      { hpt: 22 }, // Row 3 Table 1 Header
      { hpt: 20 }, // Row 4 Average
      { hpt: 20 }, // Row 5 Lowest
      { hpt: 20 }, // Row 6 Allowed Range
      { hpt: 12 }, // Row 7 Spacer
      { hpt: 24 }, // Row 8 Comparison Header
      { hpt: 22 }  // Row 9 Table 2 Header
    ];
    hotelMatrixRows.forEach(() => {
      rowHeights.push({ hpt: 20 });
    });
    ws['!rows'] = rowHeights;

    // Styling palettes exactly matching screenshot image_93bc1049.png
    const navyFillTitle = { fgColor: { rgb: "0B2545" } };
    const navyFillSection = { fgColor: { rgb: "13335A" } };
    const navyFillTable = { fgColor: { rgb: "193D6A" } };
    const mintGreenFill = { fgColor: { rgb: "D1FAE5" } }; // Light mint-green for Average & Lowest valid rates
    const whiteFill = { fgColor: { rgb: "FFFFFF" } };

    const whiteTitleFont = { name: "Calibri", sz: 14, bold: true, color: { rgb: "FFFFFF" } };
    const whiteSectionFont = { name: "Calibri", sz: 11, bold: true, color: { rgb: "FFFFFF" } };
    const whiteHeaderFont = { name: "Calibri", sz: 10, bold: true, color: { rgb: "FFFFFF" } };
    const greenSummaryFont = { name: "Calibri", sz: 10, bold: true, color: { rgb: "065F46" } }; // Dark green font
    const boldDarkFont = { name: "Calibri", sz: 10, bold: true, color: { rgb: "0F172A" } };
    const regularDarkFont = { name: "Calibri", sz: 10, color: { rgb: "334155" } };
    const mutedFont = { name: "Calibri", sz: 10, color: { rgb: "94A3B8" } };

    const thinBorder = {
      top: { style: "thin", color: { rgb: "CBD5E1" } },
      bottom: { style: "thin", color: { rgb: "CBD5E1" } },
      left: { style: "thin", color: { rgb: "CBD5E1" } },
      right: { style: "thin", color: { rgb: "CBD5E1" } }
    };

    const headerBorder = {
      top: { style: "thin", color: { rgb: "244F84" } },
      bottom: { style: "thin", color: { rgb: "244F84" } },
      left: { style: "thin", color: { rgb: "244F84" } },
      right: { style: "thin", color: { rgb: "244F84" } }
    };

    const setCell = (r: number, c: number, style: any) => {
      const ref = XLSX.utils.encode_cell({ r, c });
      if (!ws[ref]) {
        ws[ref] = { t: 's', v: '' };
      }
      ws[ref].s = style;
    };

    // Apply styles to Row 0: Title Banner
    for (let c = 0; c < totalCols; c++) {
      setCell(0, c, { fill: navyFillTitle, font: whiteTitleFont, alignment: { vertical: "center", horizontal: "left" } });
    }

    // Apply styles to Row 1: Section 1 Header
    for (let c = 0; c < totalCols; c++) {
      setCell(1, c, { fill: navyFillSection, font: whiteSectionFont, alignment: { vertical: "center", horizontal: "left" } });
    }

    // Apply styles to Row 2: Table 1 Headers
    for (let c = 0; c < totalCols; c++) {
      setCell(2, c, {
        fill: navyFillTable,
        font: whiteHeaderFont,
        alignment: { vertical: "center", horizontal: c < 2 ? "left" : "center" },
        border: headerBorder
      });
    }

    // Apply styles to Row 3: Average
    setCell(3, 0, { font: boldDarkFont, alignment: { vertical: "center", horizontal: "left" }, border: thinBorder });
    setCell(3, 1, { font: regularDarkFont, alignment: { vertical: "center", horizontal: "left" }, border: thinBorder });
    for (let c = 2; c < totalCols; c++) {
      setCell(3, c, {
        fill: mintGreenFill,
        font: greenSummaryFont,
        alignment: { vertical: "center", horizontal: "center" },
        border: thinBorder
      });
    }

    // Apply styles to Row 4: Lowest
    setCell(4, 0, { font: boldDarkFont, alignment: { vertical: "center", horizontal: "left" }, border: thinBorder });
    setCell(4, 1, { font: regularDarkFont, alignment: { vertical: "center", horizontal: "left" }, border: thinBorder });
    for (let c = 2; c < totalCols; c++) {
      setCell(4, c, {
        fill: mintGreenFill,
        font: greenSummaryFont,
        alignment: { vertical: "center", horizontal: "center" },
        border: thinBorder
      });
    }

    // Apply styles to Row 5: Allowed range
    setCell(5, 0, { font: boldDarkFont, alignment: { vertical: "center", horizontal: "left" }, border: thinBorder });
    setCell(5, 1, { font: regularDarkFont, alignment: { vertical: "center", horizontal: "left" }, border: thinBorder });
    for (let c = 2; c < totalCols; c++) {
      setCell(5, c, {
        fill: whiteFill,
        font: regularDarkFont,
        alignment: { vertical: "center", horizontal: "center" },
        border: thinBorder
      });
    }

    // Apply styles to Row 7: Section 2 Header
    for (let c = 0; c < totalCols; c++) {
      setCell(7, c, { fill: navyFillSection, font: whiteSectionFont, alignment: { vertical: "center", horizontal: "left" } });
    }

    // Apply styles to Row 8: Table 2 Headers
    for (let c = 0; c < totalCols; c++) {
      setCell(8, c, {
        fill: navyFillTable,
        font: whiteHeaderFont,
        alignment: { vertical: "center", horizontal: c === 0 ? "left" : "center" },
        border: headerBorder
      });
    }

    // Apply styles to Rows 9+: Competitor Hotels
    hotelMatrixRows.forEach((_, idx) => {
      const r = 9 + idx;
      setCell(r, 0, { font: boldDarkFont, alignment: { vertical: "center", horizontal: "left" }, border: thinBorder });
      setCell(r, 1, { font: regularDarkFont, alignment: { vertical: "center", horizontal: "center" }, border: thinBorder });
      for (let c = 2; c < totalCols; c++) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        const val = ws[cellRef]?.v;
        const isNA = !val || val === 'N/A';
        setCell(r, c, {
          font: isNA ? mutedFont : boldDarkFont,
          alignment: { vertical: "center", horizontal: "center" },
          border: thinBorder
        });
      }
    });

    return ws;
  }

  /**
   * Generates the Flow 1 Input Workbook with all 4 sheets matching user format
   */
  static exportInputWorkbook(
    hotelName: string,
    hotelRooms: HotelRoomType[],
    competitorHotels: CompetitorHotel[],
    mappedRooms: CompetitorRoomMapping[]
  ): void {
    const wb = XLSX.utils.book_new();

    // 1. hotel_data
    const wsHotel = this.buildHotelDataSheet(hotelName, hotelRooms);
    XLSX.utils.book_append_sheet(wb, wsHotel, 'hotel_data');

    // 2. competitor_hotel
    const wsCompHotel = this.buildCompetitorHotelSheet(competitorHotels);
    XLSX.utils.book_append_sheet(wb, wsCompHotel, 'competitor_hotel');

    // 3. competitor_room
    const wsCompRoom = this.buildCompetitorRoomSheet(hotelName, mappedRooms);
    XLSX.utils.book_append_sheet(wb, wsCompRoom, 'competitor_room');

    // 4. competitor_price
    const wsPrice = this.buildCompetitorPriceSheet(hotelRooms, competitorHotels, mappedRooms, undefined, hotelName);
    XLSX.utils.book_append_sheet(wb, wsPrice, 'competitor_price');

    // Save
    const cleanFilename = `${hotelName || 'hotel_mapping'}.xlsx`;
    XLSX.writeFile(wb, cleanFilename);
  }

  /**
   * Generates the Flow 2 Dated Output Workbook with all 4 sheets matching user format
   */
  static exportOutputWorkbook(
    hotelName: string,
    arrivalDate: string,
    departureDate: string,
    priceReport: PriceReportRow[],
    updatedRooms: CompetitorRoomMapping[],
    autoDownload: boolean = false,
    hotelRooms?: HotelRoomType[],
    competitorHotels?: CompetitorHotel[]
  ): { filename: string; blob: Blob } {
    const wb = XLSX.utils.book_new();

    const derivedHotelRooms: HotelRoomType[] = (hotelRooms && hotelRooms.length > 0)
      ? hotelRooms
      : priceReport.map(p => ({
          room_type: p.our_room_type,
          room_name: p.our_room_type,
          min_rate: p.our_min_rate,
          max_rate: p.our_max_rate
        }));

    const hotelMap = new Map<string, CompetitorHotel>();
    if (competitorHotels && competitorHotels.length > 0) {
      competitorHotels.forEach(h => {
        if (h.competitor_hotel_id) {
          hotelMap.set(String(h.competitor_hotel_id).trim(), h);
        }
      });
    }

    const uniqueCompIds = Array.from(new Set([
      ...(competitorHotels || []).map(h => String(h.competitor_hotel_id || '').trim()).filter(Boolean),
      ...updatedRooms.map(r => String(r.competitor_hotel_id || '').trim()).filter(Boolean)
    ]));

    const derivedCompHotels: CompetitorHotel[] = uniqueCompIds.length > 0
      ? uniqueCompIds.map(id => {
          const existing = hotelMap.get(id);
          if (existing) return existing;
          const match = updatedRooms.find(r => String(r.competitor_hotel_id).trim() === id);
          return {
            competitor_hotel_id: id,
            competitor_hotel_name: match?.competitor_hotel_name || `Hotel ${id}`
          };
        })
      : (competitorHotels && competitorHotels.length > 0 ? competitorHotels : []);

    const enrichedCompHotels = derivedCompHotels.map(h => this.enrichCompetitorHotel(h));

    // 1. hotel_data
    const wsHotel = this.buildHotelDataSheet(hotelName, derivedHotelRooms);
    XLSX.utils.book_append_sheet(wb, wsHotel, 'hotel_data');

    // 2. competitor_hotel
    const wsCompHotel = this.buildCompetitorHotelSheet(enrichedCompHotels);
    XLSX.utils.book_append_sheet(wb, wsCompHotel, 'competitor_hotel');

    // 3. competitor_room
    const wsCompRoom = this.buildCompetitorRoomSheet(hotelName, updatedRooms);
    XLSX.utils.book_append_sheet(wb, wsCompRoom, 'competitor_room');

    // 4. competitor_price
    const wsPrice = this.buildCompetitorPriceSheet(derivedHotelRooms, enrichedCompHotels, updatedRooms, priceReport, hotelName);
    XLSX.utils.book_append_sheet(wb, wsPrice, 'competitor_price');

    // Generate dated filename
    const datePrefix = new Date().toISOString().slice(0, 10);
    const filename = `${datePrefix}-${hotelName}.xlsx`;

    // Trigger download in browser ONLY if explicitly requested via download button
    if (autoDownload) {
      XLSX.writeFile(wb, filename);
    }

    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

    return { filename, blob };
  }

  /**
  /**
   * Helper to find a worksheet case-insensitively and ignoring whitespace/special characters
   */
  private static findSheet(wb: XLSX.WorkBook, ...candidatePatterns: string[]): { name: string; sheet: XLSX.WorkSheet } | null {
    const normalize = (s: string) => String(s || '').toLowerCase().replace(/[\s_\-\.]/g, '');
    for (const pat of candidatePatterns) {
      const normPat = normalize(pat);
      for (const name of wb.SheetNames) {
        const normName = normalize(name);
        if (normName === normPat || normName.includes(normPat)) {
          return { name, sheet: wb.Sheets[name] };
        }
      }
    }
    return null;
  }

  /**
   * Helper to retrieve a field value from an Excel row object case-insensitively and ignoring spaces
   */
  private static getField(row: any, ...candidates: string[]): any {
    if (!row || typeof row !== 'object') return undefined;
    for (const c of candidates) {
      if (row[c] !== undefined && row[c] !== null && String(row[c]).trim() !== '') {
        return row[c];
      }
    }
    const normalize = (s: string) => String(s || '').toLowerCase().replace(/[\s_\-\.]/g, '');
    const normCandidates = candidates.map(normalize);
    for (const [k, v] of Object.entries(row)) {
      if (v === undefined || v === null || String(v).trim() === '') continue;
      if (normCandidates.includes(normalize(k))) {
        return v;
      }
    }
    return undefined;
  }

  /**
   * Imports an existing Excel workbook (.xlsx) and parses its sheets
   */
  static async importInputWorkbook(file: File): Promise<{
    hotelName: string;
    hotelRooms: HotelRoomType[];
    competitorHotels: CompetitorHotel[];
    mappedRooms: CompetitorRoomMapping[];
  }> {
    const arrayBuffer = await file.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: 'array' });
    return this.parseWorkbookCore(wb, file.name.replace(/\.xlsx$/i, ''));
  }

  /**
   * Parses an input workbook directly from a Buffer or Uint8Array (for Node/Daemon)
   */
  static parseWorkbookFromBuffer(buffer: Buffer | Uint8Array): {
    hotelName: string;
    hotelRooms: HotelRoomType[];
    competitorHotels: CompetitorHotel[];
    mappedRooms: CompetitorRoomMapping[];
  } {
    const wb = XLSX.read(buffer, { type: 'buffer' });
    return this.parseWorkbookCore(wb, 'Hotel');
  }

  /**
   * Unified, highly resilient workbook parser that handles all casing variations,
   * flexible column headers, and auto-detects competitor_hotel_search
   */
  private static parseWorkbookCore(wb: XLSX.WorkBook, fallbackHotelName: string): {
    hotelName: string;
    hotelRooms: HotelRoomType[];
    competitorHotels: CompetitorHotel[];
    mappedRooms: CompetitorRoomMapping[];
  } {
    let hotelName = fallbackHotelName;
    let hotelRooms: HotelRoomType[] = [];
    let competitorHotels: CompetitorHotel[] = [];
    let mappedRooms: CompetitorRoomMapping[] = [];

    // 1. Parse hotel_data
    const hotelDataSheetInfo = this.findSheet(wb, 'hotel_data', 'hoteldata', 'hotel');
    if (hotelDataSheetInfo) {
      const data = XLSX.utils.sheet_to_json<any>(hotelDataSheetInfo.sheet);
      if (data.length > 0) {
        const foundName = this.getField(data[0], 'hotel_name', 'hotel', 'name');
        if (foundName) hotelName = String(foundName).trim();
      }
      hotelRooms = data.map(d => ({
        room_type: String(this.getField(d, 'room_type', 'roomtype', 'type') || '').trim(),
        room_name: String(this.getField(d, 'room_name', 'roomname', 'room_type', 'name') || '').trim(),
        min_rate: Number(this.getField(d, 'min_rate', 'minrate', 'min_price', 'minimum')) || 0,
        max_rate: Number(this.getField(d, 'max_rate', 'maxrate', 'max_price', 'maximum')) || 0
      })).filter(r => r.room_type);
    }

    // 2. Parse competitor_hotel
    let competitorHotelSheetInfo = this.findSheet(
      wb,
      'competitor_hotel',
      'competitorhotel',
      'competitor_hotels',
      'competitor',
      'competitors'
    );

    // Fallback: If not found by name, scan all sheets for competitor_hotel_search column
    if (!competitorHotelSheetInfo) {
      for (const sheetName of wb.SheetNames) {
        const rows = XLSX.utils.sheet_to_json<any>(wb.Sheets[sheetName]);
        if (rows.length > 0) {
          const hasSearchCol = rows.some(r => !!this.getField(r, 'competitor_hotel_search', 'competitor_hotel', 'hotel_search', 'search_query'));
          if (hasSearchCol) {
            competitorHotelSheetInfo = { name: sheetName, sheet: wb.Sheets[sheetName] };
            break;
          }
        }
      }
    }

    if (competitorHotelSheetInfo) {
      const data = XLSX.utils.sheet_to_json<any>(competitorHotelSheetInfo.sheet);
      competitorHotels = data.map(d => {
        const searchStr = String(
          this.getField(
            d,
            'competitor_hotel_search',
            'competitor_hotel',
            'competitor_hotel_name',
            'hotel_search',
            'search_query',
            'search_hotel',
            'hotel_name',
            'hotel',
            'query'
          ) || ''
        ).trim();

        const hotelId = String(
          this.getField(d, 'competitor_hotel_id', 'hotel_id', 'dest_id') ||
          (searchStr && /^\d+$/.test(searchStr) ? searchStr : '') ||
          ''
        ).trim();

        const hotelNameParsed = String(
          this.getField(d, 'competitor_hotel_name', 'hotel_name', 'name') ||
          (searchStr && !/^\d+$/.test(searchStr) ? searchStr : '') ||
          (hotelId ? `Hotel ${hotelId}` : '')
        ).trim();

        return {
          competitor_hotel_search: searchStr || hotelNameParsed,
          competitor_hotel_id: hotelId,
          competitor_hotel_name: hotelNameParsed || searchStr || (hotelId ? `Hotel ${hotelId}` : 'Unknown Hotel'),
          address: this.getField(d, 'address', 'hotel_address') ? String(this.getField(d, 'address', 'hotel_address')) : undefined,
          city: this.getField(d, 'city', 'location') ? String(this.getField(d, 'city', 'location')) : undefined,
          countrycode: this.getField(d, 'countrycode', 'country_code', 'country') ? String(this.getField(d, 'countrycode', 'country_code', 'country')) : undefined,
          currency_code: this.getField(d, 'currency_code', 'currency', 'currencycode')
            ? String(this.getField(d, 'currency_code', 'currency', 'currencycode')).trim().toUpperCase()
            : (this.getField(d, 'countrycode') === 'id' ? 'IDR' : 'MYR'),
          zip: this.getField(d, 'zip', 'zipcode', 'postal_code') ? String(this.getField(d, 'zip', 'zipcode', 'postal_code')) : undefined,
          review_score: this.getField(d, 'review_score', 'rating', 'score') !== undefined 
            ? Number(this.getField(d, 'review_score', 'rating', 'score')) 
            : undefined,
          url: this.getField(d, 'url', 'link') ? String(this.getField(d, 'url', 'link')) : undefined
        };
      }).filter(h => h.competitor_hotel_search || h.competitor_hotel_id || h.competitor_hotel_name);
    }

    // 3. Parse competitor_room
    const competitorRoomSheetInfo = this.findSheet(
      wb,
      'competitor_room',
      'competitorroom',
      'competitor_rooms',
      'room_mapping',
      'rooms'
    );

    if (competitorRoomSheetInfo) {
      const data = XLSX.utils.sheet_to_json<any>(competitorRoomSheetInfo.sheet);
      if ((!hotelName || hotelName === fallbackHotelName) && data.length > 0) {
        const foundOurHotel = this.getField(data[0], 'hotel_name');
        if (foundOurHotel && this.getField(data[0], 'competitor_hotel_name')) {
          hotelName = String(foundOurHotel).trim();
        }
      }
      mappedRooms = data.map(d => {
        const rawGross = this.getField(d, 'gross_amount', 'grossamount', 'price', 'rate');
        const parsedCurrency = this.getField(d, 'currency_code', 'currency', 'currencycode')
          ? String(this.getField(d, 'currency_code', 'currency', 'currencycode')).trim().toUpperCase()
          : 'MYR';
        return {
          competitor_hotel_id: String(this.getField(d, 'competitor_hotel_id', 'hotel_id') || '').trim(),
          competitor_hotel_name: String(this.getField(d, 'competitor_hotel_name', 'hotel_name') || '').trim(),
          competitor_room_id: String(this.getField(d, 'competitor_room_id', 'room_id') || '').trim(),
          competitor_room_name: String(this.getField(d, 'competitor_room_name', 'room_name', 'room') || '').trim(),
          currency_code: parsedCurrency,
          currency: parsedCurrency,
          gross_amount: rawGross !== undefined && rawGross !== '' && rawGross !== 'N/A'
            ? Number(rawGross)
            : null,
          our_room_type: String(this.getField(d, 'room_type', 'our_room_type', 'ourroomtype') || 'UNMAPPED').trim(),
          extracted_at: this.getField(d, 'extracted_at', 'extractedat', 'timestamp', 'date')
            ? String(this.getField(d, 'extracted_at', 'extractedat', 'timestamp', 'date')).trim()
            : undefined
        };
      }).filter(r => r.competitor_room_name || r.competitor_room_id);
    }

    return {
      hotelName,
      hotelRooms,
      competitorHotels,
      mappedRooms
    };
  }

  /**
   * Generates dated output workbook returning Buffer (for Node/Daemon)
   * Exact 4-sheet format matching user specifications
   */
  static generateOutputWorkbookBuffer(
    hotelName: string,
    priceReport: PriceReportRow[],
    updatedRooms: CompetitorRoomMapping[],
    hotelRooms?: HotelRoomType[],
    competitorHotels?: CompetitorHotel[]
  ): Buffer {
    const wb = XLSX.utils.book_new();

    const derivedHotelRooms: HotelRoomType[] = (hotelRooms && hotelRooms.length > 0)
      ? hotelRooms
      : priceReport.map(p => ({
          room_type: p.our_room_type,
          room_name: p.our_room_type,
          min_rate: p.our_min_rate,
          max_rate: p.our_max_rate
        }));

    const hotelMap = new Map<string, CompetitorHotel>();
    if (competitorHotels && competitorHotels.length > 0) {
      competitorHotels.forEach(h => {
        if (h.competitor_hotel_id) {
          hotelMap.set(String(h.competitor_hotel_id).trim(), h);
        }
      });
    }

    const uniqueCompIds = Array.from(new Set([
      ...(competitorHotels || []).map(h => String(h.competitor_hotel_id || '').trim()).filter(Boolean),
      ...updatedRooms.map(r => String(r.competitor_hotel_id || '').trim()).filter(Boolean)
    ]));

    const derivedCompHotels: CompetitorHotel[] = uniqueCompIds.length > 0
      ? uniqueCompIds.map(id => {
          const existing = hotelMap.get(id);
          if (existing) return existing;
          const match = updatedRooms.find(r => String(r.competitor_hotel_id).trim() === id);
          return {
            competitor_hotel_id: id,
            competitor_hotel_name: match?.competitor_hotel_name || `Hotel ${id}`
          };
        })
      : (competitorHotels && competitorHotels.length > 0 ? competitorHotels : []);

    const enrichedCompHotels = derivedCompHotels.map(h => this.enrichCompetitorHotel(h));

    // 1. hotel_data
    const wsHotel = this.buildHotelDataSheet(hotelName, derivedHotelRooms);
    XLSX.utils.book_append_sheet(wb, wsHotel, 'hotel_data');

    // 2. competitor_hotel
    const wsCompHotel = this.buildCompetitorHotelSheet(enrichedCompHotels);
    XLSX.utils.book_append_sheet(wb, wsCompHotel, 'competitor_hotel');

    // 3. competitor_room
    const wsCompRoom = this.buildCompetitorRoomSheet(hotelName, updatedRooms);
    XLSX.utils.book_append_sheet(wb, wsCompRoom, 'competitor_room');

    // 4. competitor_price
    const wsPrice = this.buildCompetitorPriceSheet(derivedHotelRooms, enrichedCompHotels, updatedRooms, priceReport, hotelName);
    XLSX.utils.book_append_sheet(wb, wsPrice, 'competitor_price');

    return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
  }
}
