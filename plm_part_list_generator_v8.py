"""
PLM Part List Generator - Version 8
부품리스트 자동화 항목 선정 방식.docx 의 규칙을 적용합니다.

[항목 선정 규칙]
  1. 트리 1레벨: 제품코드/품목명 (스킵, 자식 탐색)
  2. 2레벨 이하 적용:
     2-1. 품목명에 "별도","조립","assy","ass'y","용접" 포함 → 해당 항목 제외 (자식은 처리)
     2-2. wdww ERP 하위 자재 중 목제/목재/패널/엣지/edge 포함 또는 "-"만인 것 → 제외
     3.   "별도" 품목 하위 자재는 모두 선정
          3-1. 그 중 하위 단계가 있는 자재(모)만 선정, 자식(자)는 모의 비고에 코드 기재
  
[Excel 배열 순서 - 8번 규칙]
  - 레벨별(breadth-first): 선정된 2레벨 전부 → 3레벨 전부 → … → 별도부품 (마지막)
  - 같은 레벨 내 우선순위: wdww*** > stpg***=stcg*** > 기타
  - 동일 자재 → 1개 표기, 수량 합산

[구분 행 - 9번 규칙]
  - 기본: "-"
  - 별도 품목 하위: "별도품목"
  - 조립/assy/ass'y 품목의 자식 → "조립 {ERP}" 표기 + 병합
  - 용접 품목의 자식 → "용접 {ERP}" 표기 + 병합

[비고 행 - 10번 규칙]
  - 3레벨 이하 자재 비고: "모ERP코드: 수량" 기재 (모가 여러 개면 모두 기재)
  - 별도 3-1 규칙으로 제외된 자식들: 모의 비고에 "자재코드: 수량" 기재
"""
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import json
from collections import defaultdict
import requests
import io
from PIL import Image as PilImage
from openpyxl.drawing.image import Image as ExcelImage


# ──────────────────────────────────────────────────────────────────
#  PLM 데이터 추출 (기존 코드 그대로 유지)
# ──────────────────────────────────────────────────────────────────

def extract_grid_data(driver):
    """
    AUIGrid에서 실제 데이터를 추출합니다.
    확인된 필드명: erpCode, partName, partTypeCd, amount, new_color_0/1/...
    ※ 이 함수는 기존 plm_part_list_generator.py 와 동일합니다.
    """
    print("AUIGrid에서 데이터 추출 중 (JavaScript API)...")
    data = []

    try:
        js_code = """
        var result = {
            hasAUIGrid: false,
            gridWrapExists: false,
            myGridIDVal: null,
            method: 'none',
            error: null,
            colorColumns: [],
            allColumns: [],
            osColorMap: {},
            data: []
        };

        try {
            result.hasAUIGrid = typeof AUIGrid !== 'undefined';
            result.gridWrapExists = !!document.getElementById('grid_wrap');

            if (!result.hasAUIGrid) {
                result.error = 'AUIGrid not defined';
                return result;
            }

            function tryGetData(pid) {
                try {
                    var d = AUIGrid.getGridData(pid);
                    if (d && d.length > 0) return d;
                } catch(e) {}
                return null;
            }

            var rawData = null;
            var usedPid = null;

            if (typeof myGridID !== 'undefined' && myGridID) {
                result.myGridIDVal = myGridID;
                rawData = tryGetData(myGridID);
                if (rawData) usedPid = myGridID;
            }
            if (!rawData) { rawData = tryGetData('#grid_wrap'); if (rawData) usedPid = '#grid_wrap'; }
            if (!rawData && AUIGrid.list && AUIGrid.list.length > 0) {
                try { rawData = AUIGrid.list[0].getGridData(); if (rawData) usedPid = 'list[0]'; } catch(e) {}
            }

            if (!rawData) { result.error = 'all methods returned empty'; return result; }
            result.method = usedPid;
            result.data = rawData;

            try {
                var cols = AUIGrid.getColumnInfoList(usedPid);
                result.allColumns = [];
                if (cols) {
                    for (var c = 0; c < cols.length; c++) {
                        var df = cols[c].dataField || '';
                        var ht = cols[c].headerText || '';
                        result.allColumns.push({df: df, ht: ht});
                        if (df === 'color' || df.indexOf('color') >= 0 ||
                            df.startsWith('new_color_') ||
                            ht === '색상' || ht.indexOf('색') >= 0) {
                            result.colorColumns.push({df: df, ht: ht});
                        }
                    }
                }
            } catch(ce) { result.error = 'colInfo:' + ce.message; }

            try {
                var osColorMap = {};
                var rowMatrix = [];
                var knownColorCodes = /^(OS|WW|ML|WH|NT|BK|GR|BR|PK|NI)$/;
                var totalRows = 0;
                try { totalRows = AUIGrid.getRowCount(usedPid) || 0; } catch(re) {}
                result.totalRows = totalRows;

                for (var ri = 0; ri < totalRows; ri++) {
                    var rowItem = null;
                    try {
                        if (typeof AUIGrid.getItemByRowIndex === 'function') {
                            rowItem = AUIGrid.getItemByRowIndex(usedPid, ri);
                        }
                    } catch(re2) {}

                    if (!rowItem) continue;

                    var erp = String(rowItem.erpCode || rowItem.partNo || '').trim();
                    if (!erp) continue;

                    var colorFields = {};
                    for (var k in rowItem) {
                        if (k === 'children' || k.charAt(0) === '_') continue;
                        var v = rowItem[k];
                        if (v !== null && v !== undefined) {
                            var vs = String(v).trim();
                            if (knownColorCodes.test(vs)) {
                                colorFields[k] = vs;
                            }
                        }
                    }

                    if (ri < 10) {
                        rowMatrix.push({ri: ri, erp: erp, colorFields: colorFields,
                                         allColor: rowItem.color || ''});
                    }

                    var wwColor = String(rowItem.color || '').trim();
                    var osColor = '';

                    for (var ck in colorFields) {
                        if (ck !== 'color' && colorFields[ck] !== wwColor) {
                            osColor = colorFields[ck];
                            break;
                        }
                        if (ck !== 'color' && colorFields[ck] === wwColor) {
                            osColor = colorFields[ck];
                        }
                    }

                    if (wwColor || osColor) {
                        if (!osColorMap[erp]) osColorMap[erp] = [];
                        var isDup2 = osColorMap[erp].some(function(e) {
                            return e.os === osColor && e.ww === wwColor;
                        });
                        if (!isDup2) {
                            osColorMap[erp].push({os: osColor, ww: wwColor, ri: ri});
                        }
                    }
                }

                result.osColorMap = osColorMap;
                result.rowMatrix = rowMatrix;

            } catch(de) {
                result.domError = de.message;
            }

            try {
                var domColorTds = document.querySelectorAll('.aui-grid-header-panel td.aui-grid-default-header');
                if (domColorTds.length === 0) domColorTds = document.querySelectorAll('.aui-grid-header-panel td');
                var domColorCnt = 0;
                domColorTds.forEach(function(td) {
                    if ((td.innerText || '').trim() === '색상') domColorCnt++;
                });
                result.domColorCount = domColorCnt;
            } catch(dce) { result.domColorCount = -1; }

        } catch(e) {
            result.error = 'outer:' + e.message;
        }
        return result;
        """
        js_result = driver.execute_script(js_code)
        print(f"[JS] hasAUIGrid={js_result.get('hasAUIGrid')}, method={js_result.get('method')}, error={js_result.get('error')}")
        print(f"[JS] totalRows: {js_result.get('totalRows', 0)}")

        grid_data = js_result.get('data', [])
        color_columns = [c['df'] for c in js_result.get('colorColumns', []) if c.get('df')]

        if grid_data and len(grid_data) > 0:
            print(f"[성공] {len(grid_data)}개 최상위 아이템 추출")

            try:
                with open('_last_grid_data.json', 'w', encoding='utf-8') as f:
                    json.dump(grid_data, f, ensure_ascii=False, indent=2)
                print("[디버그] _last_grid_data.json 저장됨")
            except Exception as e:
                print(f"[경고] JSON 저장 실패: {e}")

            # 색상 컬럼 정보 포함하여 반환 (v2에서 collect_parts_v2가 처리)
            return {'grid_data': grid_data, 'color_columns': color_columns}

        else:
            print("[경고] AUIGrid API 데이터 없음. DOM 방식으로 전환...")
            return {'grid_data': [], 'color_columns': [], 'dom_data': extract_grid_data_from_dom(driver)}

    except Exception as e:
        print(f"[에러] JavaScript 추출 실패: {e}")
        print("DOM 방식으로 전환...")
        return {'grid_data': [], 'color_columns': [], 'dom_data': extract_grid_data_from_dom(driver)}


def extract_grid_data_from_dom(driver):
    """
    Fallback: DOM에서 직접 데이터 추출
    ※ 이 함수는 기존 plm_part_list_generator.py 와 동일합니다.
    """
    print("DOM에서 데이터 추출 중 (data-field 기반)...")
    data = []

    try:
        time.sleep(2)

        header_cells = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'aui-grid-header-panel')]//th | "
            "//div[contains(@class,'aui-grid-header')]//th"
        )
        col_order = []
        for idx, hdr in enumerate(header_cells):
            df = (hdr.get_attribute('data-field') or
                  hdr.get_attribute('data-column-field') or
                  hdr.get_attribute('column-field') or '')
            col_order.append((idx, df))

        field_idx = {df: i for i, df in col_order if df}

        rows = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'aui-grid-body-panel')]//table//tr"
        )
        print(f"[DOM] 바디 행 수: {len(rows)}")

        def get_cell_val(cells, field_name):
            if field_name in field_idx:
                idx2 = field_idx[field_name]
                if idx2 < len(cells):
                    c = cells[idx2]
                    v = c.text.strip()
                    if not v:
                        v = (c.get_attribute('data-value') or c.get_attribute('title') or '')
                    return v
            return ''

        for row in rows:
            cells = row.find_elements(By.TAG_NAME, 'td')
            if not cells:
                continue

            if field_idx:
                erp_code  = get_cell_val(cells, 'erpCode')
                part_name = get_cell_val(cells, 'partName')
                qty       = get_cell_val(cells, 'amount')
                part_type = get_cell_val(cells, 'partTypeCd')
                colors = []
                for df, cidx in field_idx.items():
                    if df.startswith('color_') or df.startswith('new_color_'):
                        if cidx < len(cells):
                            colors.append(cells[cidx].text.strip())
                color_val = colors
            else:
                erp_code  = cells[1].text.strip() if len(cells) > 1 else ''
                part_name = cells[3].text.strip() if len(cells) > 3 else ''
                qty       = cells[12].text.strip() if len(cells) > 12 else ''
                part_type = cells[9].text.strip() if len(cells) > 9 else ''
                color_val = []

            if erp_code or part_name:
                data.append({
                    "erpCode":    erp_code,
                    "partName":   part_name,
                    "amount":     qty,
                    "partTypeCd": part_type,
                    "color":      color_val,
                    "_$depth":    2,
                    "children":   [],
                })

        print(f"[DOM] 추출된 데이터: {len(data)}개")

    except Exception as e:
        print(f"[에러] DOM 추출 실패: {e}")
        import traceback
        traceback.print_exc()

    return data


# ──────────────────────────────────────────────────────────────────
#  V2 항목 선정 로직
# ──────────────────────────────────────────────────────────────────

def _erp(item):
    return str(item.get('erpCode') or item.get('partNo') or '').strip()

def _name(item):
    return str(item.get('partName') or '').strip()

def _qty(item):
    try:
        return float(item.get('amount') or 1)
    except (ValueError, TypeError):
        return 1.0

def _depth(item):
    try:
        return int(item.get('_$depth', 0))
    except (ValueError, TypeError):
        return 0

def _is_baldo(item):
    """별도 항목: ETPJ로 시작하거나 품목명에 '별도' 포함"""
    return _erp(item).upper().startswith('ETPJ') or '별도' in _name(item)

def _is_assembly(item):
    """조립/용접 키워드 포함 여부 (v5 변경 규칙: 빌드, 웰드, build, weld)"""
    name = _name(item).lower()
    return ('빌드' in name or 'build' in name or '웰드' in name or 'weld' in name)

def _assembly_label(item):
    """구분 레이블: '조립' 또는 '용접' (v5 변경 표기)"""
    name = _name(item).lower()
    if '웰드' in name or 'weld' in name:
        return f"[용접] {_name(item)}"
    return f"[조립] {_name(item)}"

def _erp_priority(erp_code):
    """같은 레벨 내 정렬 우선순위: wdww=0, stpg/stcg=1, 기타=2"""
    e = erp_code.upper()
    if e.startswith('WDWW'):
        return 0
    if e.startswith('STPG') or e.startswith('STCG'):
        return 1
    return 2

def _is_wdww_excluded(item):
    """
    wdww 하위 자재 중 제외 대상 (규칙 2-2):
    - 품목명에 목제/목재/패널/엣지/edge 포함
    - 품목명이 '-' 또는 비어있음
    ※ 자신의 ERP가 wdww여도 품목명 조건에 해당하면 제외
    ※ 단, '스티커'가 포함된 품목은 예외적으로 제외하지 않음 (예: 목재수분 주의스티커)
    """
    name = _name(item)
    EXCLUDE_KW = ('목제', '목재', '패널', '엣지', 'edge')
    if not name or not name.replace('-', '').strip():
        return True
    
    # 스티커 등은 키워드가 포함되어 있어도 제외하지 않음
    if '스티커' in name:
        return False
        
    return any(kw.lower() in name.lower() for kw in EXCLUDE_KW)


def _round_qty(val):
    """부동소수점 오차 제거 후 정수면 정수, 아니면 소수점 4자리로 반올림"""
    rounded = round(val, 4)
    if rounded == int(rounded):
        return int(rounded)
    return rounded


def _extract_colors(item, color_columns):
    """color_columns 필드에서 색상 값 추출 → 리스트 반환"""
    c_list = []
    if color_columns:
        for c_df in color_columns:
            v = str(item.get(c_df, '')).strip()
            c_list.append(v if v not in ('None', 'null', '') else '')
    else:
        # fallback: color_ 또는 new_color_ 로 시작하는 필드 모두 수집
        color_vals = []
        for k, v in item.items():
            if (k.startswith('color_') or k.startswith('new_color_')) and v:
                vs = str(v).strip()
                if vs and vs not in ('', 'None', 'null'):
                    color_vals.append(vs)
        c = item.get('color', '')
        if c and str(c).strip() not in ('', 'None', 'null'):
            color_vals.insert(0, str(c).strip())
        c_list = color_vals if color_vals else ['']
    return c_list


def make_part_dict(item, qty, div1, div2, div3, remarks=None, is_baldo=False, color_columns=None):
    return {
        'erp_code':   _erp(item),
        'part_name':  _name(item),
        'part_type':  str(item.get('partTypeCd') or item.get('partType') or '').strip(),
        'thumb':      str(item.get('thumb') or '').replace('\\', '/').strip(),
        'qty':        qty,
        'colors':     _extract_colors(item, color_columns or []),
        'depth':      _depth(item),
        'div_level1': div1,
        'div_level2': div2,
        'div_level3': div3,
        'remarks':    list(remarks) if remarks else [],
        '_is_baldo':  is_baldo,
    }


def extract_wdww_sub_items_recursive(item, parent_qty, parent_erp, top_div1='-', color_columns=None, collector=None):
    """
    [규칙 2-2] WDWW 항목 하위의 비제외 자재(케이싱, 케이싱스크류, 스티커, 네임플레이트 등) 재귀 추출 및 병합:
    - 목재, 목제, 패널, 엣지, edge 또는 '-' 항목은 해당 자재 및 하위 전체를 제외.
    - 케이싱, 케이싱스크류, 스티커, 네임플레이트 등 비제외 부속 자재는 수집하여
      동일 자재코드별 수량 합산 및 비고(부모코드: 수량) 누적.
    """
    if collector is None:
        collector = {}

    for ch in item.get('children', []):
        if _is_wdww_excluded(ch):
            continue

        c_erp = _erp(ch)
        c_qty = _round_qty(_qty(ch) * parent_qty)
        rem_entry = f"{parent_erp}: {c_qty}" if parent_erp else ""

        key = c_erp.upper() if c_erp else _name(ch)
        if key in collector:
            collector[key]['qty'] = _round_qty(collector[key]['qty'] + c_qty)
            if rem_entry and rem_entry not in collector[key]['remarks']:
                collector[key]['remarks'].append(rem_entry)
        else:
            p_dict = make_part_dict(
                ch, c_qty, top_div1, '-', '-', remarks=[rem_entry] if rem_entry else [],
                is_baldo=False, color_columns=color_columns
            )
            collector[key] = p_dict

        # 하위 자식이 더 있을 경우 재귀 탐색
        extract_wdww_sub_items_recursive(ch, c_qty, parent_erp, top_div1, color_columns, collector)

    return collector


def _collect_assembly_group(asm_item, parent_qty, top_asm_label, color_columns=None):
    """
    빌드/웰드 모자재 1개 및 그 하위 자재들을 재귀적으로 수집 및 정렬:
    [정렬 순서]:
      1. 모자재 본인 (is_asm_head = True, div_level2 = sub_asm_label)
      2. 목재 단품 (WDWW)
      3. 철제 단품 (STPG, STCG)
      4. 기타 단품 (볼트, 사출, 하드웨어 등)
      5. 하위 [빌드] 세트 (재귀)
      6. 하위 [웰드] 세트 (재귀)
    """
    if color_columns is None:
        color_columns = []

    sub_asm_label = f"{_assembly_label(asm_item)} {_erp(asm_item)}"
    is_weld = ('웰드' in _name(asm_item).lower() or 'weld' in _name(asm_item).lower())
    sub_color_type = 'weld' if is_weld else 'build'
    ach_qty = _round_qty(_qty(asm_item) * parent_qty)

    item_dict = make_part_dict(
        asm_item, _round_qty(_qty(asm_item)), top_asm_label, sub_asm_label, None, remarks=[], is_baldo=False, color_columns=color_columns
    )
    item_dict['is_asm_head'] = True
    item_dict['asm_color_type'] = sub_color_type

    sub_wdww = []
    sub_steel = []
    sub_hardware = []
    sub_build = []
    sub_weld = []

    for a_sub in asm_item.get('children', []):
        a_sub_qty = _round_qty(_qty(a_sub) * ach_qty)
        a_sub_erp = _erp(a_sub).upper()

        if _is_assembly(a_sub):
            nested_group, nested_is_weld = _collect_assembly_group(
                a_sub, ach_qty, top_asm_label, color_columns
            )
            if nested_is_weld:
                sub_weld.extend(nested_group)
            else:
                sub_build.extend(nested_group)
        else:
            c_dict = make_part_dict(
                a_sub, a_sub_qty, top_asm_label, f"__CHILD__{sub_asm_label}", '-', remarks=[], is_baldo=False, color_columns=color_columns
            )
            c_dict['asm_color_type'] = sub_color_type
            if a_sub_erp.startswith('WDWW'):
                sub_wdww.append(c_dict)
            elif a_sub_erp.startswith('STPG') or a_sub_erp.startswith('STCG'):
                sub_steel.append(c_dict)
            else:
                sub_hardware.append(c_dict)

    sub_group = [item_dict] + sub_wdww + sub_steel + sub_hardware + sub_build + sub_weld
    return sub_group, is_weld


def collect_parts_v8(grid_data, color_columns=None):
    """
    PLM 트리에서 항목 선정 - V8 변경 규칙 적용:
    - 품목명에 '빌드(build)' 또는 '웰드(weld)'가 포함된 경우:
      - 메인 부품리스트(main_parts)에는 해당 조립/용접 자재 본인만 1개 열로 등재.
        - 비고란에는 직속 하위 자식 목록('자재코드 : 수량') 기재.
        - 하위 자식들은 메인 리스트에서 전개하지 않음.
      - 조립_용접 시트(assy_parts)에는 해당 조립/용접 자재의 하위 구성 자재들을 전개하여 수집.
        - 상단 3개 구분 행으로 조립/용접 계층 표기.
    """
    if color_columns is None:
        color_columns = []

    true_roots = [item for item in grid_data if _depth(item) <= 1]
    if not true_roots:
        true_roots = [grid_data[0]] if grid_data else []

    main_normal = []
    main_baldo  = []
    main_wdww_sub_collector = {}

    bucket_wdww     = []
    bucket_steel    = []
    bucket_hardware = []
    bucket_build    = []
    bucket_weld     = []
    assy_baldo      = []
    assy_wdww_sub_collector = {}

    for root in true_roots:
        for ch in root.get('children', []):
            my_qty = _round_qty(_qty(ch))
            erp = _erp(ch)
            children = ch.get('children', [])

            if _is_baldo(ch):
                baldo_lv2 = f"별도품목 {erp}" if erp else "별도품목"
                for bch in children:
                    b_qty = _round_qty(_qty(bch) * my_qty)
                    b_children = bch.get('children', [])

                    if _is_assembly(bch):
                        # [V8 규칙] 빌드/웰드 품목은 본인만 메인 시트에 등재 (별도품목 구분)
                        main_baldo.append(make_part_dict(
                            bch, b_qty, baldo_lv2, baldo_lv2, '-', remarks=[], is_baldo=True, color_columns=color_columns
                        ))

                        # 조립_용접 시트에 하위 자재들 수집 (Group 6) - 5단계 버킷 정렬
                        b_asm_label = f"[조립]{_name(bch)} {_erp(bch)}"
                        b_wdww = []
                        b_steel = []
                        b_hardware = []
                        b_build = []
                        b_weld = []
                        for sub in b_children:
                            sub_qty = _round_qty(_qty(sub) * b_qty)
                            sub_erp = _erp(sub).upper()
                            p_dict = make_part_dict(
                                sub, sub_qty, baldo_lv2, b_asm_label, '-', remarks=[], is_baldo=True, color_columns=color_columns
                            )
                            if _is_assembly(sub):
                                is_w = ('웰드' in _name(sub).lower() or 'weld' in _name(sub).lower())
                                if is_w:
                                    b_weld.append(p_dict)
                                else:
                                    b_build.append(p_dict)
                            elif sub_erp.startswith('WDWW'):
                                b_wdww.append(p_dict)
                            elif sub_erp.startswith('STPG') or sub_erp.startswith('STCG'):
                                b_steel.append(p_dict)
                            else:
                                b_hardware.append(p_dict)
                        assy_baldo.extend(b_wdww + b_steel + b_hardware + b_build + b_weld)
                    else:
                        # 일반 자재
                        if b_children:
                            rem = [f"{_erp(sc)}: {_round_qty(_qty(sc))}" for sc in b_children if _erp(sc)]
                            main_baldo.append(make_part_dict(
                                bch, b_qty, baldo_lv2, baldo_lv2, '-', remarks=rem, is_baldo=True, color_columns=color_columns
                            ))
                        else:
                            main_baldo.append(make_part_dict(
                                bch, b_qty, baldo_lv2, baldo_lv2, '-', remarks=[], is_baldo=True, color_columns=color_columns
                            ))

            elif _is_assembly(ch):
                # [V8 규칙] 빌드/웰드 품목 본인을 메인 시트에 등재!
                # 비고란에 직속 자식 목록 (코드 : 수량) 기재
                rem_lines = [f"{_erp(c)} : {_round_qty(_qty(c))}" for c in children if _erp(c)]
                main_normal.append(make_part_dict(
                    ch, my_qty, '-', '-', '-', remarks=rem_lines, is_baldo=False, color_columns=color_columns
                ))

                # 조립_용접 시트에 하위 전개 항목들 수집: 6단계 버킷 정렬
                top_asm_label = f"{_assembly_label(ch)} {_erp(ch)}"
                for ach in children:
                    ach_qty = _round_qty(_qty(ach) * my_qty)

                    if _is_assembly(ach):
                        sub_group, is_weld = _collect_assembly_group(
                            ach, my_qty, top_asm_label, color_columns=color_columns
                        )
                        if is_weld:
                            # Group 5: 일반 [용접] 세트
                            bucket_weld.extend(sub_group)
                        else:
                            # Group 4: 일반 [빌드] 세트
                            bucket_build.extend(sub_group)
                    else:
                        ach_dict = make_part_dict(
                            ach, ach_qty, top_asm_label, '-', '-', remarks=[], is_baldo=False, color_columns=color_columns
                        )
                        ach_erp = _erp(ach).upper()
                        if ach_erp.startswith('WDWW'):
                            # Group 1: 목재 단품 자재
                            bucket_wdww.append(ach_dict)
                            # [규칙 2-2] WDWW 하위 자재 재귀 탐색 (목재/엣지 등 제외, 케이싱/스크류/스티커/네임플레이트 등 수집)
                            extract_wdww_sub_items_recursive(
                                ach, ach_qty, _erp(ach), top_div1=top_asm_label, color_columns=color_columns, collector=assy_wdww_sub_collector
                            )
                        elif ach_erp.startswith('STPG') or ach_erp.startswith('STCG'):
                            # Group 2: 철제 단품 자재
                            bucket_steel.append(ach_dict)
                            if ach.get('children'):
                                extract_wdww_sub_items_recursive(
                                    ach, ach_qty, _erp(ach), top_div1=top_asm_label, color_columns=color_columns, collector=assy_wdww_sub_collector
                                )
                        else:
                            # Group 3: 기타 단품 자재 (모듈, 사출 등)
                            bucket_hardware.append(ach_dict)
                            # 하위 자재(스티커, 부속품 등) 재귀 탐색
                            if ach.get('children'):
                                extract_wdww_sub_items_recursive(
                                    ach, ach_qty, _erp(ach), top_div1=top_asm_label, color_columns=color_columns, collector=assy_wdww_sub_collector
                                )

            else:
                ch_dict = make_part_dict(
                    ch, my_qty, '-', '-', '-', remarks=[], is_baldo=False, color_columns=color_columns
                )
                main_normal.append(ch_dict)

                # 메인 시트 2레벨 단품에 하위 자재가 있는 경우 재귀 탐색 (WDWW 목재 부속, 모듈 스티커 등)
                if ch.get('children'):
                    extract_wdww_sub_items_recursive(
                        ch, my_qty, erp, top_div1='-', color_columns=color_columns, collector=main_wdww_sub_collector
                    )

    # 1. 메인 시트: 2레벨 자재들 뒤에 WDWW 하위 부품(케이싱, 케이싱스크류, 스티커, 네임플레이트 등) 나열 ("따로 빼서 뒤에 나열")
    main_normal.extend(main_wdww_sub_collector.values())
    main_parts = main_normal + main_baldo

    # 2. 조립_용접 시트: WDWW 하위 부품들을 bucket_hardware 뒤에 병합 나열
    for hw_key, hw_item in assy_wdww_sub_collector.items():
        found = False
        for existing in bucket_hardware:
            if existing['erp_code'].upper() == hw_key:
                existing['qty'] = _round_qty(existing['qty'] + hw_item['qty'])
                for rem in hw_item.get('remarks', []):
                    if rem not in existing['remarks']:
                        existing['remarks'].append(rem)
                found = True
                break
        if not found:
            bucket_hardware.append(hw_item)

    assy_parts = (
        bucket_wdww +
        bucket_steel +
        bucket_hardware +
        bucket_build +
        bucket_weld +
        assy_baldo
    )
    return main_parts, assy_parts


def collect_parts_v2(grid_data, color_columns=None):
    m, _ = collect_parts_v8(grid_data, color_columns)
    return m


def write_partlist_table(ws, parts, start_row, max_colors=1, num_div_rows=1, driver=None):
    """
    공통 테이블 작성기:
    - num_div_rows=1: 메인 시트용 (구분 1줄, 26.1pt)
    - num_div_rows=3: 조립_용접 시트용 (구분 3줄: 25.5pt, 25.5pt, 32.1pt)
    - 최대 12열 단위 청크 분할 및 개행 배치
    - 썸네일 이미지 삽입 및 텍스트 자동 크기 조정 지원
    """
    f_bold    = Font(bold=True, size=10)
    fill_hdr  = PatternFill("solid", fgColor="D9D9D9")
    fill_dash = PatternFill("solid", fgColor="BDD7EE")
    fill_assy = PatternFill("solid", fgColor="E2EFDA")
    fill_bal  = PatternFill("solid", fgColor="FFD966")
    fill_build_lv2 = PatternFill("solid", fgColor="FCD5B4")
    fill_build_lv3 = PatternFill("solid", fgColor="FDE9D9")
    fill_weld_lv2  = PatternFill("solid", fgColor="CCC0DA")
    fill_weld_lv3  = PatternFill("solid", fgColor="E4DFEC")
    ac = Alignment(horizontal="center", vertical="center", wrap_text=True)
    at = Alignment(horizontal="left",   vertical="top",    wrap_text=True)
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'),  bottom=Side(style='thin'))

    LABEL_COL  = 2
    DATA_START = 3
    CHUNK_SIZE = 12
    n = len(parts)

    current_row_offset = start_row

    for chunk_start in range(0, n, CHUNK_SIZE):
        chunk_parts = parts[chunk_start : chunk_start + CHUNK_SIZE]
        num_chunk = len(chunk_parts)

        if num_div_rows == 1:
            ROW_DIV   = current_row_offset + 1
            ROW_NAME  = current_row_offset + 2
            ROW_TYPE  = current_row_offset + 3
            ROW_COL_S = current_row_offset + 4
            ROW_COL_E = ROW_COL_S + max_colors - 1
            ROW_QTY   = ROW_COL_E + 1
            ROW_ERP   = ROW_QTY   + 1
            ROW_REM   = ROW_ERP   + 1
            div_row_keys = [(ROW_DIV, 'div_level2')]
        else:
            ROW_DIV1  = current_row_offset + 1
            ROW_DIV2  = current_row_offset + 2
            ROW_DIV3  = current_row_offset + 3
            ROW_NAME  = current_row_offset + 4
            ROW_TYPE  = current_row_offset + 5
            ROW_COL_S = current_row_offset + 6
            ROW_COL_E = ROW_COL_S + max_colors - 1
            ROW_QTY   = ROW_COL_E + 1
            ROW_ERP   = ROW_QTY   + 1
            ROW_REM   = ROW_ERP   + 1
            div_row_keys = [
                (ROW_DIV1, 'div_level1'),
                (ROW_DIV2, 'div_level2'),
                (ROW_DIV3, 'div_level3')
            ]

        # Col B 행 레이블
        for r_idx, _ in div_row_keys:
            c = ws.cell(row=r_idx, column=LABEL_COL, value="구분")
            c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin

        c = ws.cell(row=ROW_NAME, column=LABEL_COL, value="부품명")
        c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin

        c = ws.cell(row=ROW_TYPE, column=LABEL_COL, value="형태")
        c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin

        for i in range(max_colors):
            c = ws.cell(row=ROW_COL_S + i, column=LABEL_COL, value="색상" if i == 0 else "")
            c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin
        if max_colors > 1:
            ws.merge_cells(start_row=ROW_COL_S, start_column=LABEL_COL,
                           end_row=ROW_COL_E,   end_column=LABEL_COL)

        c = ws.cell(row=ROW_QTY, column=LABEL_COL, value="수량")
        c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin

        c = ws.cell(row=ROW_ERP, column=LABEL_COL, value="자재코드")
        c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin

        c = ws.cell(row=ROW_REM, column=LABEL_COL, value="비고")
        c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin

        # 구분 행 병합 및 기입
        if num_div_rows == 1:
            col = DATA_START
            while col < DATA_START + num_chunk:
                idx = col - DATA_START
                cur = chunk_parts[idx].get('div_level2', '-')
                end_col = col
                if cur not in ('-', ''):
                    while end_col + 1 < DATA_START + num_chunk and chunk_parts[end_col+1-DATA_START].get('div_level2') == cur:
                        end_col += 1
                div_fill = fill_dash if cur in ('-', '') else (fill_bal if '별도품목' in cur else fill_assy)
                for c_border in range(col, end_col + 1):
                    c_cell = ws.cell(row=ROW_DIV, column=c_border, value=cur if c_border == col else None)
                    c_cell.fill = div_fill; c_cell.border = thin; c_cell.alignment = ac
                    if c_border == col:
                        if cur and len(str(cur)) > 22:
                            c_cell.font = Font(bold=True, size=8.5)
                        elif cur and len(str(cur)) > 14:
                            c_cell.font = Font(bold=True, size=9)
                        else:
                            c_cell.font = f_bold
                if end_col > col:
                    ws.merge_cells(start_row=ROW_DIV, start_column=col,
                                   end_row=ROW_DIV,   end_column=end_col)
                col = end_col + 1
        else:
            # 1) ROW_DIV1: 대분류 (예: [조립] 익스텐션 다리 빌드... / 별도품목 ETPJ...)
            col = DATA_START
            while col < DATA_START + num_chunk:
                idx = col - DATA_START
                cur = chunk_parts[idx].get('div_level1', '-')
                end_col = col
                if cur not in ('-', ''):
                    while end_col + 1 < DATA_START + num_chunk and chunk_parts[end_col+1-DATA_START].get('div_level1') == cur:
                        end_col += 1
                div_fill = fill_bal if '별도품목' in cur else (fill_dash if cur in ('-', '') else fill_assy)
                for c_border in range(col, end_col + 1):
                    c_cell = ws.cell(row=ROW_DIV1, column=c_border, value=cur if c_border == col else None)
                    c_cell.fill = div_fill; c_cell.border = thin; c_cell.alignment = ac
                    if c_border == col:
                        c_cell.font = Font(bold=True, size=8.5) if (cur and len(str(cur)) > 20) else f_bold
                if end_col > col:
                    ws.merge_cells(start_row=ROW_DIV1, start_column=col,
                                   end_row=ROW_DIV1,   end_column=end_col)
                col = end_col + 1

            # 2) ROW_DIV2 & ROW_DIV3
            col = DATA_START
            while col < DATA_START + num_chunk:
                idx = col - DATA_START
                p = chunk_parts[idx]
                is_asm_head = p.get('is_asm_head', False)
                cur2 = p.get('div_level2', '-')
                ctype = p.get('asm_color_type')

                if is_asm_head:
                    # 모자재 헤드: 세로 2행 병합 (ROW_DIV2 ~ ROW_DIV3)
                    head_fill = fill_weld_lv2 if ctype == 'weld' else fill_build_lv2
                    c2_cell = ws.cell(row=ROW_DIV2, column=col, value=cur2)
                    c2_cell.fill = head_fill; c2_cell.border = thin; c2_cell.alignment = ac
                    c2_cell.font = Font(bold=True, size=8.0) if (cur2 and len(str(cur2)) > 18) else f_bold
                    c3_cell = ws.cell(row=ROW_DIV3, column=col, value=None)
                    c3_cell.fill = head_fill; c3_cell.border = thin; c3_cell.alignment = ac
                    ws.merge_cells(start_row=ROW_DIV2, start_column=col,
                                   end_row=ROW_DIV3,   end_column=col)
                    col += 1
                elif str(cur2).startswith('__CHILD__'):
                    # 모자재 하위 자식들의 ROW_DIV2 영역: 같은 자식들끼리 가로 병합, 값은 빈 문자열
                    end_col = col
                    while end_col + 1 < DATA_START + num_chunk and chunk_parts[end_col+1-DATA_START].get('div_level2') == cur2:
                        end_col += 1
                    child_fill = fill_weld_lv2 if ctype == 'weld' else fill_build_lv2
                    child_fill_lv3 = fill_weld_lv3 if ctype == 'weld' else fill_build_lv3
                    for c_border in range(col, end_col + 1):
                        c2_cell = ws.cell(row=ROW_DIV2, column=c_border, value="" if c_border == col else None)
                        c2_cell.fill = child_fill; c2_cell.border = thin; c2_cell.alignment = ac
                        # ROW_DIV3 영역: 자식 각각에 '-' 기입 및 연한 틴트 색상
                        c3_cell = ws.cell(row=ROW_DIV3, column=c_border, value="-")
                        c3_cell.fill = child_fill_lv3; c3_cell.border = thin; c3_cell.alignment = ac
                        c3_cell.font = f_bold
                    if end_col > col:
                        ws.merge_cells(start_row=ROW_DIV2, start_column=col,
                                       end_row=ROW_DIV2,   end_column=end_col)
                    col = end_col + 1
                elif cur2 in ('-', '', None):
                    # 일반 자재: 가로 병합 없이 각각 '-'
                    c2_cell = ws.cell(row=ROW_DIV2, column=col, value="-")
                    c2_cell.fill = fill_dash; c2_cell.border = thin; c2_cell.alignment = ac; c2_cell.font = f_bold
                    c3_cell = ws.cell(row=ROW_DIV3, column=col, value="-")
                    c3_cell.fill = fill_dash; c3_cell.border = thin; c3_cell.alignment = ac; c3_cell.font = f_bold
                    col += 1
                else:
                    # 일반 조립 블록 (예: 별도품목 하위 조립 품목의 자식들)
                    end_col = col
                    while end_col + 1 < DATA_START + num_chunk and chunk_parts[end_col+1-DATA_START].get('div_level2') == cur2:
                        end_col += 1
                    for c_border in range(col, end_col + 1):
                        c2_cell = ws.cell(row=ROW_DIV2, column=c_border, value=cur2 if c_border == col else None)
                        c2_cell.fill = fill_assy; c2_cell.border = thin; c2_cell.alignment = ac
                        if c_border == col:
                            c2_cell.font = Font(bold=True, size=8.5) if len(str(cur2)) > 20 else f_bold
                        c3_cell = ws.cell(row=ROW_DIV3, column=c_border, value="-")
                        c3_cell.fill = fill_assy; c3_cell.border = thin; c3_cell.alignment = ac; c3_cell.font = f_bold
                    if end_col > col:
                        ws.merge_cells(start_row=ROW_DIV2, start_column=col,
                                       end_row=ROW_DIV2,   end_column=end_col)
                    col = end_col + 1

        # 각 부품 데이터 기입
        for i, p in enumerate(chunk_parts):
            col = DATA_START + i
            colors = p.get('colors', [])
            if isinstance(colors, str):
                colors = [c.strip() for c in colors.split(',') if c.strip()]
            while len(colors) < max_colors:
                colors.append('')

            # 부품명
            c = ws.cell(row=ROW_NAME, column=col, value=p.get('part_name') or None)
            c.alignment = ac; c.border = thin

            # 형태 (이미지 다운로드 및 삽입)
            cell = ws.cell(row=ROW_TYPE, column=col)
            thumb_url = p.get('thumb')
            img_inserted = False
            if thumb_url and 'productview_openin' not in thumb_url and driver:
                import html
                import base64
                thumb_url = html.unescape(thumb_url)
                if thumb_url.startswith('/'):
                    thumb_url = "https://plm.fursys.com" + thumb_url
                try:
                    driver.set_script_timeout(15)
                    script = """
                    var uri = arguments[0];
                    var callback = arguments[1];
                    fetch(uri)
                      .then(res => res.blob())
                      .then(blob => {
                          var reader = new FileReader();
                          reader.onloadend = function() { callback(reader.result); }
                          reader.readAsDataURL(blob);
                      })
                      .catch(err => callback('error:' + err.toString()));
                    """
                    b64_data = driver.execute_async_script(script, thumb_url)
                    if b64_data and str(b64_data).startswith('data:image'):
                        header, encoded = b64_data.split(",", 1)
                        img_data = io.BytesIO(base64.b64decode(encoded))
                        pil_img = PilImage.open(img_data)
                        orig_w, orig_h = pil_img.size
                        MAX_W = 140
                        MAX_H = 90
                        scale = min(MAX_W / float(orig_w), MAX_H / float(orig_h))
                        new_w = max(1, int(orig_w * scale))
                        new_h = max(1, int(orig_h * scale))
                        xl_img = ExcelImage(img_data)
                        xl_img.width = new_w
                        xl_img.height = new_h
                        ws.add_image(xl_img, cell.coordinate)
                        cell.value = ""
                        img_inserted = True
                except Exception as e:
                    print(f"  [이미지 오류] {e}")
            if not img_inserted:
                cell.value = p.get('part_type') or None
            cell.alignment = ac; cell.border = thin

            # 색상
            for ci, cv in enumerate(colors):
                c = ws.cell(row=ROW_COL_S+ci, column=col, value=cv or None)
                c.alignment = ac; c.border = thin

            # 수량
            c = ws.cell(row=ROW_QTY, column=col, value=p.get('qty') or None)
            c.alignment = ac; c.border = thin

            # 자재코드
            c = ws.cell(row=ROW_ERP, column=col, value=p.get('erp_code') or None)
            c.alignment = ac; c.border = thin

            # 비고
            rem_list = p.get('remarks', [])
            rem_str = "\n".join(str(x) for x in rem_list)
            c = ws.cell(row=ROW_REM, column=col, value=rem_str or None)
            c.alignment = at; c.border = thin

        # 행 높이 설정
        if num_div_rows == 1:
            ws.row_dimensions[ROW_DIV].height = 26.1
        else:
            ws.row_dimensions[ROW_DIV1].height = 25.5
            ws.row_dimensions[ROW_DIV2].height = 25.5
            ws.row_dimensions[ROW_DIV3].height = 32.1

        ws.row_dimensions[ROW_NAME].height = 49.5
        ws.row_dimensions[ROW_TYPE].height = 72.0
        for ci in range(max_colors):
            ws.row_dimensions[ROW_COL_S+ci].height = 21.95
        ws.row_dimensions[ROW_QTY].height = 21.95
        ws.row_dimensions[ROW_ERP].height = 21.95

        # 비고 행 높이: 기본 66.0pt 유지하되, 4줄 초과 시 모든 줄이 보이도록 동적 확장
        max_rem_lines = 0
        for p in chunk_parts:
            rem_list = p.get('remarks', [])
            if isinstance(rem_list, list):
                cnt = sum(len(str(x).split('\n')) for x in rem_list)
            else:
                cnt = len(str(rem_list).split('\n'))
            if cnt > max_rem_lines:
                max_rem_lines = cnt

        if max_rem_lines > 4:
            ws.row_dimensions[ROW_REM].height = max(66.0, max_rem_lines * 18.5 + 10.0)
        else:
            ws.row_dimensions[ROW_REM].height = 66.0

        current_row_offset = ROW_REM + 1

    return current_row_offset


def generate_part_list_excel_v3(parts, output_file, product_code="", product_name="", edge_info=None, cookies=None, driver=None, version_str="R000", assy_parts=None):
    """
    Excel 생성 (V8):
      - 메인 시트([단품코드]): 빌드/웰드 모자재 본인 등재, 비고에 하위 자식 표기, 1개 구분 행
      - 조립_용접 시트: 조립/용접 하위 자재 상세 전개 (3개 구분 행)
      - 출력 페이지: 2페이지 A4 규격 자동 분할 (1페이지: 메인+색상표, 2페이지: 조립/용접)
    """
    print(f"[v8] Excel 생성: {output_file}, 메인 부품 수: {len(parts)}, 조립부품 수: {len(assy_parts) if assy_parts else 0}")
    
    import sys
    import copy
    template_path = None
    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        cand_ext = os.path.join(exe_dir, 'template.xlsx')
        cand_int = os.path.join(getattr(sys, '_MEIPASS', exe_dir), 'template.xlsx')
        if os.path.exists(cand_ext):
            template_path = cand_ext
        elif os.path.exists(cand_int):
            template_path = cand_int
    else:
        cand_dev = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template.xlsx')
        if os.path.exists(cand_dev):
            template_path = cand_dev
        
    has_assy = bool(assy_parts and len(assy_parts) > 0)

    if template_path and os.path.exists(template_path):
        print(f"[정보] 템플릿 로드: {template_path}")
        wb = openpyxl.load_workbook(template_path)
        if "출력 페이지" in wb.sheetnames:
            ws_out = wb["출력 페이지"]
            ws_out["W28"] = version_str
            ws_out["P28"] = product_name
            ws_out["P30"] = product_code
            
            # 1페이지 높이 (2~27행) 고정 (27.95pt)
            for r in range(2, 28):
                ws_out.row_dimensions[r].height = 27.95

            if has_assy:
                ws_out["W31"] = "1/2"
                ws_out.row_dimensions[34].height = 16.5
                # 2페이지 데이터 영역 높이 (35~60행) 고정
                for r in range(35, 61):
                    ws_out.row_dimensions[r].height = 27.95

                # 2페이지 타이틀 블록 복사 (28~33행 -> 61~66행)
                for r in range(28, 34):
                    tr = r + 33
                    ws_out.row_dimensions[tr].height = ws_out.row_dimensions[r].height or 16.5
                    for c_idx in range(1, 26):
                        sc = ws_out.cell(r, c_idx)
                        dc = ws_out.cell(tr, c_idx)
                        dc.value = sc.value
                        dc.font = copy.copy(sc.font)
                        dc.alignment = copy.copy(sc.alignment)
                        dc.border = copy.copy(sc.border)
                        dc.fill = copy.copy(sc.fill)

                for m in list(ws_out.merged_cells.ranges):
                    if 28 <= m.min_row <= 33:
                        ws_out.merge_cells(
                            start_row=m.min_row + 33, start_column=m.min_col,
                            end_row=m.max_row + 33,   end_column=m.max_col
                        )

                ws_out["P61"] = "조립/용접 자재"
                ws_out["P63"] = product_code
                ws_out["W61"] = version_str
                ws_out["W64"] = "2/2"
            else:
                ws_out["W31"] = "1/1"
                
        ws_title = (product_code[:31] if product_code else "PartList")
        ws = wb.create_sheet(title=ws_title)
    else:
        print(f"[경고] template.xlsx 파일을 찾을 수 없습니다! 기본 시트로 생성됩니다.")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = (product_code[:31] if product_code else "PartList")

    f_title  = Font(bold=True, size=11)
    f_code   = Font(bold=True, size=12)
    fill_hdr  = PatternFill("solid", fgColor="D9D9D9")
    ac = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'),  bottom=Side(style='thin'))

    LABEL_COL  = 2
    DATA_START = 3

    # 메인 시트 기본 설정
    ws.column_dimensions['B'].width = 14
    for i in range(max(12, len(parts))):
        ws.column_dimensions[get_column_letter(DATA_START+i)].width = 18
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 49.5

    # Row 2, 3: 단품코드/단품명
    if product_code:
        c = ws.cell(row=2, column=LABEL_COL, value="단품코드"); c.font = f_title; c.alignment = ac; c.fill = fill_hdr; c.border = thin
        c = ws.cell(row=2, column=DATA_START, value=product_code); c.font = f_code; c.alignment = ac; c.border = thin
    if product_name:
        c = ws.cell(row=3, column=LABEL_COL, value="단품명"); c.font = f_title; c.alignment = ac; c.fill = fill_hdr; c.border = thin
        c = ws.cell(row=3, column=DATA_START, value=product_name); c.alignment = al; c.border = thin

    max_colors = max((len(p.get('colors', [])) for p in parts), default=1)
    max_colors = max(max_colors, 1)

    # 메인 시트 테이블 작성 (구분 행 1줄)
    write_partlist_table(ws, parts, start_row=3, max_colors=max_colors, num_div_rows=1, driver=driver)

    # 조립_용접 시트 작성 (있는 경우)
    if has_assy:
        ws_assy = wb.create_sheet(title="조립_용접")
        ws_assy.column_dimensions['B'].width = 14
        for i in range(max(12, len(assy_parts))):
            ws_assy.column_dimensions[get_column_letter(DATA_START+i)].width = 18

        max_colors_assy = max((len(p.get('colors', [])) for p in assy_parts), default=1)
        max_colors_assy = max(max_colors_assy, 1)
        write_partlist_table(ws_assy, assy_parts, start_row=1, max_colors=max_colors_assy, num_div_rows=3, driver=driver)

    # 색상표 시트 작성 (있는 경우)
    if edge_info:
        ws2 = wb.create_sheet(title="색상표")
        f_bold = Font(bold=True, size=10)
        if edge_info.get('type') == 'simple' or 'type' not in edge_info:
            ws2.merge_cells(start_row=2, start_column=2, end_row=3, end_column=3)
            c = ws2.cell(row=2, column=2, value="색상표"); c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin
            ws2.cell(row=2, column=3).border = thin
            ws2.cell(row=3, column=2).border = thin
            ws2.cell(row=3, column=3).border = thin

            c = ws2.cell(row=4, column=2, value="표면재"); c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin
            c = ws2.cell(row=4, column=3, value="엣지"); c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin
            
            surf_colors = edge_info.get('surface', [])
            edge_colors = edge_info.get('edge', [])
            
            for i in range(max(len(surf_colors), len(edge_colors))):
                r = 5 + i
                sc = surf_colors[i] if i < len(surf_colors) else ""
                ec = edge_colors[i] if i < len(edge_colors) else ""
                c1 = ws2.cell(row=r, column=2, value=sc); c1.alignment = ac; c1.border = thin
                c2 = ws2.cell(row=r, column=3, value=ec); c2.alignment = ac; c2.border = thin

            ws2.column_dimensions['B'].width = 14
            ws2.column_dimensions['C'].width = 14
            
        elif edge_info.get('type') == 'complex':
            ws2.merge_cells(start_row=2, start_column=2, end_row=3, end_column=5)
            c = ws2.cell(row=2, column=2, value="색상표"); c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin
            for r in [2, 3]:
                for c_idx in range(2, 6):
                    ws2.cell(row=r, column=c_idx).border = thin
                    
            headers = ["표면재", "부품명", "자재코드", "엣지"]
            for idx, h in enumerate(headers):
                c = ws2.cell(row=4, column=2+idx, value=h)
                c.font = f_bold; c.fill = fill_hdr; c.alignment = ac; c.border = thin
                
            r = 5
            for row_data in edge_info['rows']:
                sc = row_data['surface']
                ec = row_data['edge']
                erps = row_data['erps']
                names = row_data['names']
                count = row_data.get('count', 1)
                
                c1 = ws2.cell(row=r, column=2, value=sc); c1.alignment = ac; c1.border = thin
                c2 = ws2.cell(row=r, column=3, value=names); c2.alignment = ac; c2.border = thin
                c3 = ws2.cell(row=r, column=4, value=erps); c3.alignment = ac; c3.border = thin
                c4 = ws2.cell(row=r, column=5, value=ec); c4.alignment = ac; c4.border = thin
                
                if count > 1:
                    ws2.row_dimensions[r].height = (18 * count) + 5
                else:
                    ws2.row_dimensions[r].height = 20
                r += 1
                
            start_r = 5
            while start_r < r:
                end_r = start_r
                cur_val = ws2.cell(row=start_r, column=2).value
                while end_r + 1 < r and ws2.cell(row=end_r + 1, column=2).value == cur_val:
                    end_r += 1
                if end_r > start_r:
                    ws2.merge_cells(start_row=start_r, start_column=2, end_row=end_r, end_column=2)
                start_r = end_r + 1
                
            ws2.column_dimensions['B'].width = 14
            ws2.column_dimensions['C'].width = 40
            ws2.column_dimensions['D'].width = 20
            ws2.column_dimensions['E'].width = 14

    CHUNK_SIZE = 12
    n_first = min(len(parts), CHUNK_SIZE)
    pdf_col = DATA_START + n_first
    pdf_row = 7

    wb.save(output_file)
    
    apply_linked_pictures(output_file, ws.title, has_assy=has_assy)
    
    print(f"[v8 완료] 저장: {output_file} | 메인시트: {ws.title} | 부품수: {len(parts)} | 조립부품: {len(assy_parts) if assy_parts else 0}")
    return (output_file, pdf_row, pdf_col)
def apply_linked_pictures(excel_path, data_sheet_name, has_assy=False):
    print(f"  [v8] '출력 페이지' 연결된 그림 복사 및 PDF 저장 시작 (has_assy={has_assy})...")
    import time
    xl = None
    wb = None
    try:
        import win32com.client
        import pythoncom
        pythoncom.CoInitialize()
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(os.path.abspath(excel_path))
        
        ws_out = None
        for sh in wb.Sheets:
            if sh.Name == "출력 페이지":
                ws_out = sh
                break
                
        if not ws_out:
            print("  [경고] '출력 페이지' 시트를 찾을 수 없어 그림 복사를 건너뜁니다.")
            return
            
        target_range = ws_out.Range("B2:Y26")
        target_w = target_range.Width
        target_h = target_range.Height
        cell_b2 = ws_out.Range("B2")
        target_top = cell_b2.Top
        target_left = cell_b2.Left
        
        # 1페이지 메인 데이터 시트 탐색
        ws_data = None
        for sh in wb.Sheets:
            if sh.Name == data_sheet_name:
                ws_data = sh
                break
        if not ws_data:
            for sh in wb.Sheets:
                if sh.Name not in ["출력 페이지", "색상표", "조립_용접"]:
                    ws_data = sh
                    break

        pic1 = None
        if ws_data:
            used_range = ws_data.UsedRange
            for attempt in range(3):
                try:
                    ws_data.Activate()
                    used_range.Copy()
                    time.sleep(0.3)
                    ws_out.Activate()
                    try:
                        cell_b2.Select()
                    except Exception:
                        pass
                    ws_out.Pictures().Paste(Link=True)
                    pic1 = ws_out.Shapes(ws_out.Shapes.Count)
                    print(f"  → 부품리스트 연결된 그림(Live Link) 복사 성공 (시도 {attempt+1})")
                    break
                except Exception as e_link:
                    print(f"  [재시도 {attempt+1}] 부품리스트 Link 복사 지연/실패: {e_link}")
                    time.sleep(0.4)

            if not pic1:
                try:
                    print("  → [대체 모드] 부품리스트 정적 그림(CopyPicture) 복사 시도...")
                    ws_data.Activate()
                    used_range.CopyPicture(1, 2)
                    time.sleep(0.3)
                    ws_out.Activate()
                    try:
                        cell_b2.Select()
                    except Exception:
                        pass
                    ws_out.Pictures().Paste()
                    pic1 = ws_out.Shapes(ws_out.Shapes.Count)
                    print("  → 부품리스트 정적 그림 복사 성공")
                except Exception as e_cp:
                    print(f"  [경고] 부품리스트 그림 복사 전체 실패: {e_cp}")

            if pic1:
                try:
                    pic1.Top = target_top + 3
                    pic1.Left = target_left + 3
                    pic1.LockAspectRatio = -1
                    try:
                        pic1.Fill.Visible = True
                        pic1.Fill.Solid()
                        pic1.Fill.ForeColor.RGB = 16777215  # White
                    except Exception:
                        pass
                except Exception as e_p1:
                    print(f"  [경고] 부품리스트 기본 위치 지정 실패: {e_p1}")

        # 색상표 그림 복사 (1페이지 우측)
        has_color = False
        pic2 = None
        ws_color = None
        for sh in wb.Sheets:
            if sh.Name == "색상표":
                has_color = True
                ws_color = sh
                break

        if has_color and ws_color:
            used_color = ws_color.UsedRange
            for attempt in range(3):
                try:
                    ws_color.Activate()
                    used_color.Copy()
                    time.sleep(0.3)
                    ws_out.Activate()
                    try:
                        cell_b2.Select()
                    except Exception:
                        pass
                    ws_out.Pictures().Paste(Link=True)
                    pic2 = ws_out.Shapes(ws_out.Shapes.Count)
                    print(f"  → 색상표 연결된 그림 복사 성공 (시도 {attempt+1})")
                    break
                except Exception as e_c_link:
                    print(f"  [재시도 {attempt+1}] 색상표 Link 복사 지연/실패: {e_c_link}")
                    time.sleep(0.4)

            if not pic2:
                try:
                    print("  → [대체 모드] 색상표 정적 그림(CopyPicture) 복사 시도...")
                    ws_color.Activate()
                    used_color.CopyPicture(1, 2)
                    time.sleep(0.3)
                    ws_out.Activate()
                    try:
                        cell_b2.Select()
                    except Exception:
                        pass
                    ws_out.Pictures().Paste()
                    pic2 = ws_out.Shapes(ws_out.Shapes.Count)
                    print("  → 색상표 정적 그림 복사 성공")
                except Exception as e_c_cp:
                    print(f"  [경고] 색상표 그림 복사 전체 실패: {e_c_cp}")

            if pic2:
                try:
                    pic2.LockAspectRatio = -1
                    try:
                        pic2.Fill.Visible = True
                        pic2.Fill.Solid()
                        pic2.Fill.ForeColor.RGB = 16777215
                    except Exception:
                        pass
                except Exception as e_p2:
                    print(f"  [경고] 색상표 기본 위치 지정 실패: {e_p2}")

        # 1페이지 크기 비율 조정
        try:
            if pic1 and has_color and pic2:
                orig_w1, orig_h1 = pic1.Width, pic1.Height
                orig_w2, orig_h2 = pic2.Width, pic2.Height
                pic2.Width = orig_w2 * 0.8
                padding = 20
                remaining_w = target_w - pic2.Width - padding
                scale1 = 1.0
                if orig_h1 > target_h:
                    scale1 = target_h / orig_h1
                if (orig_w1 * scale1) > remaining_w:
                    scale1 = remaining_w / orig_w1
                pic1.Width = orig_w1 * scale1
                pic2.Top = pic1.Top
                pic2.Left = pic1.Left + pic1.Width + padding
            elif pic1:
                orig_w1, orig_h1 = pic1.Width, pic1.Height
                scale = 1.0
                if orig_h1 > target_h:
                    scale = target_h / orig_h1
                if (orig_w1 * scale) > target_w:
                    scale = target_w / orig_w1
                pic1.Width = orig_w1 * scale
        except Exception as e_scale:
            print(f"  [경고] 1페이지 크기 비율 조정 실패: {e_scale}")

        # 2페이지: 조립_용접 시트 그림 복사
        ws_assy = None
        for sh in wb.Sheets:
            if sh.Name == "조립_용접":
                ws_assy = sh
                break

        if ws_assy:
            cell_b35 = ws_out.Range("B35")
            target_range_p2 = ws_out.Range("B35:Y60")
            target_w2 = target_range_p2.Width
            target_h2 = target_range_p2.Height
            used_assy = ws_assy.UsedRange

            pic_assy = None
            for attempt in range(3):
                try:
                    ws_assy.Activate()
                    used_assy.Copy()
                    time.sleep(0.3)
                    ws_out.Activate()
                    try:
                        cell_b35.Select()
                    except Exception:
                        pass
                    ws_out.Pictures().Paste(Link=True)
                    pic_assy = ws_out.Shapes(ws_out.Shapes.Count)
                    print(f"  → 조립_용접 연결된 그림(Live Link) 복사 성공 (시도 {attempt+1})")
                    break
                except Exception as e_a_link:
                    print(f"  [재시도 {attempt+1}] 조립_용접 Link 복사 지연/실패: {e_a_link}")
                    time.sleep(0.4)

            if not pic_assy:
                try:
                    print("  → [대체 모드] 조립_용접 정적 그림(CopyPicture) 복사 시도...")
                    ws_assy.Activate()
                    used_assy.CopyPicture(1, 2)
                    time.sleep(0.3)
                    ws_out.Activate()
                    try:
                        cell_b35.Select()
                    except Exception:
                        pass
                    ws_out.Pictures().Paste()
                    pic_assy = ws_out.Shapes(ws_out.Shapes.Count)
                    print("  → 조립_용접 정적 그림 복사 성공")
                except Exception as e_a_cp:
                    print(f"  [경고] 조립_용접 그림 복사 전체 실패: {e_a_cp}")

            if pic_assy:
                try:
                    pic_assy.Top = cell_b35.Top + 3
                    pic_assy.Left = cell_b35.Left + 3
                    pic_assy.LockAspectRatio = -1
                    try:
                        pic_assy.Fill.Visible = True
                        pic_assy.Fill.Solid()
                        pic_assy.Fill.ForeColor.RGB = 16777215
                    except Exception:
                        pass

                    orig_wa, orig_ha = pic_assy.Width, pic_assy.Height
                    scale_a = 1.0
                    if orig_ha > target_h2:
                        scale_a = target_h2 / orig_ha
                    if (orig_wa * scale_a) > target_w2:
                        scale_a = target_w2 / orig_wa
                    pic_assy.Width = orig_wa * scale_a
                except Exception as e_pa:
                    print(f"  [경고] 조립_용접 위치/크기 조정 실패: {e_pa}")

            # 2페이지 구분선 및 인쇄 영역 설정
            try:
                ws_out.HPageBreaks.Add(ws_out.Range("A34"))
                print("  → 2페이지 구분선 (A34) 설정 완료")
            except Exception as e_hpb:
                print(f"  [경고] HPageBreak 설정 실패: {e_hpb}")

            ws_out.PageSetup.PrintArea = "$A$1:$Y$66"
        else:
            ws_out.PageSetup.PrintArea = "$A$1:$Y$33"

        ws_out.PageSetup.Orientation = 2  # Landscape
        ws_out.PageSetup.PaperSize = 9    # A4
        ws_out.PageSetup.Zoom = 48

        # 엑셀 변경사항 저장
        try:
            wb.Save()
            print("  → 엑셀 통합 문서 저장 완료")
        except Exception as e_save:
            print(f"  [경고] 엑셀 저장 실패: {e_save}")

        # PDF 자동 내보내기
        try:
            pdf_path = os.path.splitext(os.path.abspath(excel_path))[0] + ".pdf"
            ws_out.ExportAsFixedFormat(0, pdf_path)
            print(f"  → PDF 저장 완료: {os.path.basename(pdf_path)}")
        except Exception as e_pdf:
            print(f"  [경고] '출력 페이지' PDF 저장 실패: {e_pdf}")
            try:
                wb.ExportAsFixedFormat(0, pdf_path)
                print(f"  → (대체) 전체 통합문서 PDF 저장 완료: {os.path.basename(pdf_path)}")
            except Exception as e_pdf2:
                print(f"  [경고] 전체 통합문서 PDF 저장도 실패: {e_pdf2}")

        print("  → '출력 페이지' 세팅 완료!")

    except Exception as e:
        print(f"  [경고] '출력 페이지' 처리 중 예기치 않은 오류: {e}")
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=True)
            if xl:
                xl.Quit()
        except Exception:
            pass
# ──────────────────────────────────────────────────────────────────
#  v7: PDF OLE 삽입 / 조립도 없음 표시
# ──────────────────────────────────────────────────────────────────

def attach_pdf_to_excel_ole(excel_path, pdf_path, pdf_row, pdf_col):
    """
    win32com을 사용하여 PDF를 Excel에 OLE 개체로 삽입.
    (삽입 > 개체 > 파일로부터 만들기 방식, 아이콘으로 표시)
    ※ OLE 삽입이 실패할 경우, PDF의 각 페이지를 이미지로 변환하여 '조립도' 시트에 임베딩하고 메인 시트에 바로가기 생성.

    excel_path : 저장된 Excel 파일 절대경로
    pdf_path   : 다운로드된 PDF 파일 절대경로
    pdf_row    : OLE를 삽입할 행 번호 (1-indexed)
    pdf_col    : OLE를 삽입할 열 번호 (1-indexed)
    Returns True on success, False on failure.
    """
    print(f"[v8] PDF OLE 개체 삽입 시도: {os.path.basename(pdf_path)} → (행{pdf_row}, 열{pdf_col})")
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("  [오류] pywin32 패키지가 없습니다. 'pip install pywin32' 후 재시도하세요.")
        return False

    _os = os
    pythoncom.CoInitialize()
    xl = None
    wb = None
    tmp_pdf_path = None

    # ── OLE 삽입 전 PDF를 ASCII 전용 임시 경로로 복사 ──────────────
    # 파일명에 괄호/공백/한글 등이 있으면 Excel COM 오류(-2146827284) 발생
    import shutil
    pdf_dir = _os.path.dirname(_os.path.abspath(pdf_path))
    tmp_pdf_path = _os.path.join(pdf_dir, "tmp_ole_insert.pdf")
    original_pdf_basename = _os.path.basename(pdf_path)

    try:
        shutil.copy2(_os.path.abspath(pdf_path), tmp_pdf_path)
        print(f"  -> 임시 PDF 복사: {tmp_pdf_path}")
    except Exception as e:
        print(f"  [경고] 임시 PDF 복사 실패 ({e}), 원본 경로로 시도")
        tmp_pdf_path = _os.path.abspath(pdf_path)

    try:
        # 독립적인 프로세스로 실행
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        excel_abs = _os.path.abspath(excel_path)

        wb = xl.Workbooks.Open(excel_abs)
        ws = wb.Sheets(1)
        ws.Activate()   # 시트 활성화 (OLE 삽입에 필요)
        wb.Activate()   # 워크북 활성화

        # 대상 셀 위치 및 크기 (포인트 단위)
        cell   = ws.Cells(pdf_row, pdf_col)
        left   = cell.Left
        top    = cell.Top
        width  = ws.Columns(pdf_col).ColumnWidth * 7.5
        height = ws.Rows(pdf_row).RowHeight

        ole_ok = False

        # ── 방법 1: OLEObjects().Add ──────────────────────────────────
        try:
            ole_obj = ws.OLEObjects().Add(
                Filename=tmp_pdf_path,
                Link=False,
                DisplayAsIcon=True,
                Left=left,
                Top=top,
                Width=width,
                Height=height
            )
            try:
                ole_obj.Object.Caption = original_pdf_basename
            except Exception:
                pass
            print("  -> OLE 임베딩 성공 (방법1)")
            ole_ok = True
        except Exception as ole_e1:
            print(f"  [경고] OLE 방법1 실패: {ole_e1}")

        # ── 방법 2: Package 클래스타입으로 OLE 삽입 (Adobe 없어도 가능) ──────
        if not ole_ok:
            try:
                ole_obj = ws.OLEObjects().Add(
                    ClassType="Package",
                    Filename=tmp_pdf_path,
                    Link=False,
                    DisplayAsIcon=True,
                    Left=left,
                    Top=top,
                    Width=width,
                    Height=height
                )
                print("  -> OLE 임베딩 성공 (방법2: Package)")
                ole_ok = True
            except Exception as ole_e2:
                print(f"  [경고] OLE 방법2 실패: {ole_e2}")

        # OLE 삽입 성공 또는 실패 여부에 관계없이 COM 자원 먼저 정리 및 저장
        if wb:
            wb.Save()
            wb.Close(SaveChanges=True)
            wb = None
        if xl:
            xl.Quit()
            xl = None

        if ole_ok:
            print(f"  → OLE 개체 삽입 완료: {excel_abs}")
            return True

        # ── 방법 3: OLE 실패 시 이미지 임베딩 백업 (하이퍼링크 방식 대체) ────
        print("  [경고] OLE 개체 삽입이 실패했습니다. '조립도' 이미지 시트 삽입 백업 로직을 시작합니다.")
        import fitz
        from openpyxl.drawing.image import Image as OPImage
        from openpyxl.styles import Font as OPFont, Alignment as OPAlignment, Border as OPBorder, Side as OPSide

        doc = fitz.open(pdf_path)
        wb2 = openpyxl.load_workbook(excel_path)

        # 기존 조립도 시트가 존재하면 삭제 후 재성생
        if "조립도" in wb2.sheetnames:
            del wb2["조립도"]
        ws_draw = wb2.create_sheet(title="조립도")

        temp_images = []
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=150)
            img_filename = f"temp_page_{i}.png"
            img_path = _os.path.join(pdf_dir, img_filename)
            pix.save(img_path)
            temp_images.append(img_path)

            # openpyxl로 조립도 이미지 삽입 (페이지당 50행 간격 배치)
            img = OPImage(img_path)
            ws_draw.add_image(img, f"A{1 + i * 50}")
            print(f"  -> 조립도 {i+1}페이지 이미지 삽입 완료")

        # 메인 시트의 OLE 셀에 조립도 시트로 이동하는 내부 바로가기 수식 작성
        ws_main = wb2.worksheets[0]
        cell_range = ws_main.cell(row=pdf_row, column=pdf_col)
        cell_range.value = '=HYPERLINK("#\'조립도\'!A1", "[조립도 시트 이동]")'

        # 셀 스타일링 (볼드, 파란색, 밑줄, 가운데 정렬, 테두리)
        thin_side = OPSide(style='thin')
        cell_range.font = OPFont(bold=True, size=9, color="0000FF", underline="single")
        cell_range.alignment = OPAlignment(horizontal="center", vertical="center", wrap_text=True)
        cell_range.border = OPBorder(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        wb2.save(excel_path)
        print("  → '조립도' 이미지 시트 백업 생성 및 내부 하이퍼링크 생성 성공!")

        # 임시 이미지 파일들 제거
        for img_path in temp_images:
            try:
                if _os.path.exists(img_path):
                    _os.remove(img_path)
            except Exception:
                pass

        return True

    except Exception as e:
        print(f"  [오류] Excel 처리 실패: {e}")
        import traceback
        traceback.print_exc()
        return False
    finally:
        try:
            if wb:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if xl:
                xl.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        # 임시 PDF 삭제
        try:
            if tmp_pdf_path and tmp_pdf_path != _os.path.abspath(pdf_path) and _os.path.exists(tmp_pdf_path):
                _os.remove(tmp_pdf_path)
                print(f"  -> 임시 PDF 삭제 완료")
        except Exception:
            pass


def write_no_assembly_note(excel_path, pdf_row, pdf_col):
    """
    조립도 PDF가 없을 경우, 해당 셀에 '조립도 없음' 텍스트를 삽입 (openpyxl).
    """
    print(f"[v8] '조립도 없음' 표시: (행{pdf_row}, 열{pdf_col})")
    try:
        import os as _os
        wb2 = openpyxl.load_workbook(excel_path)
        ws2 = wb2.active

        _thin = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'),  bottom=Side(style='thin')
        )
        cell = ws2.cell(row=pdf_row, column=pdf_col, value="조립도 없음")
        cell.font      = Font(bold=True, size=9, color="FF0000")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = _thin

        wb2.save(excel_path)
        print(f"  → '조립도 없음' 표시 완료")
    except Exception as e:
        print(f"  [경고] '조립도 없음' 표시 실패: {e}")


def process_and_generate(driver_result, output_file, product_code="", product_name="", driver=None, version_str="R000"):
    """
    extract_grid_data() 결과를 받아 collect_parts_v2 → generate_part_list_excel_v2 실행.
    driver_result: extract_grid_data() 반환값 (dict)
    """
    grid_data     = driver_result.get('grid_data', [])
    color_columns = driver_result.get('color_columns', [])
    dom_data      = driver_result.get('dom_data', [])
    cookies       = driver_result.get('cookies', {})

    edge_info = None

    if grid_data:
        wdww_edge_map = {}
        erp_name_map = {}
        
        def traverse_edge_info(nodes, current_wdww=None):
            for node in nodes:
                erp = _erp(node).upper()
                name = _name(node).strip()
                if erp:
                    erp_name_map[erp] = name
                
                my_wdww = current_wdww
                if erp.startswith('WDWW'):
                    my_wdww = erp
                    if my_wdww not in wdww_edge_map:
                        wdww_edge_map[my_wdww] = None
                
                if my_wdww and ('엣지' in name or 'edge' in name):
                    if not wdww_edge_map[my_wdww]:
                        wdww_edge_map[my_wdww] = _extract_colors(node, color_columns)
                
                traverse_edge_info(node.get('children', []), my_wdww)

        traverse_edge_info(grid_data)
        
        wdww_edge_map = {k: v for k, v in wdww_edge_map.items() if v}
        
        if wdww_edge_map:
            true_roots = [it for it in grid_data if int(it.get('_$depth', 0)) <= 1]
            if true_roots:
                root_cols = _extract_colors(true_roots[0], color_columns)
                
                first_edge_colors = list(wdww_edge_map.values())[0]
                all_same = all(colors == first_edge_colors for colors in wdww_edge_map.values())
                
                if all_same:
                    edge_info = {
                        'type': 'simple',
                        'surface': root_cols,
                        'edge': first_edge_colors
                    }
                else:
                    max_c = max(len(root_cols), max(len(c) for c in wdww_edge_map.values()))
                    grouped_rows = []
                    for i in range(max_c):
                        surface_color = root_cols[i] if i < len(root_cols) else ""
                        edge_to_erps = defaultdict(list)
                        for erp, edge_cols in wdww_edge_map.items():
                            ec = edge_cols[i] if i < len(edge_cols) else ""
                            edge_to_erps[ec].append(erp)
                            
                        for ec, erps in edge_to_erps.items():
                            sorted_erps = sorted(erps)
                            grouped_rows.append({
                                'surface': surface_color,
                                'names': "\n".join(erp_name_map.get(e, '') for e in sorted_erps),
                                'erps': "\n".join(sorted_erps),
                                'edge': ec,
                                'count': len(sorted_erps)
                            })
                    edge_info = {
                        'type': 'complex',
                        'rows': grouped_rows
                    }

        # 1레벨(최상위) 항목에서 단품코드/단품명 추출
        true_roots = [it for it in grid_data if _depth(it) <= 1]
        if true_roots:
            root = true_roots[0]
            product_code = _erp(root) or product_code
            product_name = _name(root) or product_name

        main_parts, assy_parts = collect_parts_v8(grid_data, color_columns)
        parts = main_parts
    elif dom_data:
        # DOM fallback: dom_data는 flat list, 그대로 변환
        print("[v2] DOM 데이터로 변환 중...")
        parts = []
        for d in dom_data:
            if not (_erp(d) or _name(d)):
                continue
            colors = d.get('color', [])
            if isinstance(colors, str):
                colors = [colors] if colors else []
            parts.append({
                'erp_code':   _erp(d),
                'part_name':  _name(d),
                'part_type':  str(d.get('partTypeCd') or '').strip(),
                'qty':        str(d.get('amount', '')),
                'colors':     colors,
                'depth':      2,
                'div_level2': '-',
                'div_level3': '-',
                'remarks':    [],
                '_is_baldo':  False,
            })
    else:
        print("[v2] 데이터 없음")
        return None

    if not parts:
        print("[v3] 선정된 항목 없음")
        return None

    return generate_part_list_excel_v3(parts, output_file, product_code, product_name, edge_info, cookies, driver=driver, version_str=version_str, assy_parts=assy_parts)
