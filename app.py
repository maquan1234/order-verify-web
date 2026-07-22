# -*- coding: utf-8 -*-
"""电商审单核对Web应用 - Flask后端"""
import os
import tempfile
import time
from flask import Flask, request, jsonify, send_file, render_template
from verifier import (
    load_products, save_products, init_products_from_xlsx, build_barcode_to_info,
    build_code_to_info, match_columns_from_workbook, verify_orders, export_results_to_xlsx
)
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 云部署时通过环境变量 DATA_DIR 指向持久化卷（如 Railway/Render 挂载的磁盘），
# 保证货品数据与核对结果在实例重启后不丢失。本地运行时默认使用程序所在目录。
DATA_DIR = os.environ.get('DATA_DIR', BASE_DIR)
os.makedirs(DATA_DIR, exist_ok=True)

PRODUCTS_JSON = os.path.join(DATA_DIR, 'products.json')
# 桌面货品表仅作为本地首次初始化的可选来源，云端不存在时自动忽略
PRODUCTS_XLSX = os.environ.get('PRODUCTS_XLSX', r'C:\Users\EDY\Desktop\货品信息表.xlsx')
OUTPUT_DIR = os.path.join(DATA_DIR, 'outputs')
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__, static_folder='static', template_folder='templates')
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max upload

# ========== 初始化货品信息 ==========
def ensure_products():
    if not os.path.exists(PRODUCTS_JSON):
        # 首次部署：若持久卷为空，但代码仓库内自带 products.json（含初始货品），则种子化过去
        seed = os.path.join(BASE_DIR, 'products.json')
        if DATA_DIR != BASE_DIR and os.path.exists(seed):
            import shutil
            shutil.copyfile(seed, PRODUCTS_JSON)
        elif os.path.exists(PRODUCTS_XLSX):
            init_products_from_xlsx(PRODUCTS_XLSX, PRODUCTS_JSON)
        else:
            save_products(PRODUCTS_JSON, [])

ensure_products()

# ========== 页面路由 ==========
@app.route('/')
def index():
    return render_template('index.html')

# ========== 货品信息API ==========
@app.route('/api/products', methods=['GET'])
def get_products():
    products = load_products(PRODUCTS_JSON)
    return jsonify(products)

@app.route('/api/products', methods=['POST'])
def add_product():
    data = request.json
    products = load_products(PRODUCTS_JSON)
    new_id = max([p['id'] for p in products], default=0) + 1
    data['id'] = new_id
    products.append(data)
    save_products(PRODUCTS_JSON, products)
    return jsonify({'success': True, 'product': data})

@app.route('/api/products/<int:pid>', methods=['PUT'])
def update_product(pid):
    data = request.json
    products = load_products(PRODUCTS_JSON)
    for p in products:
        if p['id'] == pid:
            p['name'] = data.get('name', p['name'])
            p['code'] = data.get('code', p['code'])
            p['barcode'] = data.get('barcode', p['barcode'])
            p['common'] = data.get('common', p['common'])
            save_products(PRODUCTS_JSON, products)
            return jsonify({'success': True, 'product': p})
    return jsonify({'success': False, 'error': '货品不存在'}), 404

@app.route('/api/products/<int:pid>', methods=['DELETE'])
def delete_product(pid):
    products = load_products(PRODUCTS_JSON)
    products = [p for p in products if p['id'] != pid]
    save_products(PRODUCTS_JSON, products)
    return jsonify({'success': True})

@app.route('/api/products/import', methods=['POST'])
def import_products():
    """从xlsx文件导入货品信息（覆盖）"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '未上传文件'}), 400
    f = request.files['file']
    tmp_path = os.path.join(tempfile.gettempdir(), f'prod_{int(time.time())}.xlsx')
    f.save(tmp_path)
    try:
        products = init_products_from_xlsx(tmp_path, PRODUCTS_JSON)
        return jsonify({'success': True, 'count': len(products)})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

# ========== 审单核对API ==========
@app.route('/api/verify', methods=['POST'])
def verify_upload():
    """上传销售单xlsx，自动核对，返回结果摘要和下载ID"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '未上传文件'}), 400
    f = request.files['file']
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'success': False, 'error': '请上传xlsx文件'}), 400

    tmp_path = os.path.join(tempfile.gettempdir(), f'upload_{int(time.time())}.xlsx')
    f.save(tmp_path)

    try:
        wb_raw = load_workbook(tmp_path, data_only=True)

        # 检查sheet结构，支持不同sheet名称
        sheet_names = wb_raw.sheetnames
        if len(sheet_names) < 2:
            wb_raw.close()
            return jsonify({'success': False, 'error': '文件需至少包含2个工作表（订单信息+货品明细）'}), 400

        # 如果sheet名称不是sheet0/sheet1，自动重命名
        if 'sheet0' not in sheet_names or 'sheet1' not in sheet_names:
            # 尝试根据内容判断哪个是sheet0（订单汇总）哪个是sheet1（货品明细）
            from verifier import _read_headers, _find_col_index
            ws_a = wb_raw[sheet_names[0]]
            ws_b = wb_raw[sheet_names[1]]
            headers_a = _read_headers(ws_a)
            headers_b = _read_headers(ws_b)
            # sheet1有货品编号或条码，sheet0有合并备注或网店订单号
            has_product_b = _find_col_index(headers_b, ['货品编号', '条码']) >= 0
            has_product_a = _find_col_index(headers_a, ['货品编号', '条码']) >= 0
            if has_product_b and not has_product_a:
                wb_raw[sheet_names[0]].title = 'sheet0'
                wb_raw[sheet_names[1]].title = 'sheet1'
            elif has_product_a and not has_product_b:
                wb_raw[sheet_names[0]].title = 'sheet1'
                wb_raw[sheet_names[1]].title = 'sheet0'
            else:
                # 默认第一个是sheet0，第二个是sheet1
                wb_raw[sheet_names[0]].title = 'sheet0'
                wb_raw[sheet_names[1]].title = 'sheet1'

        matched_rows = match_columns_from_workbook(wb_raw)
        wb_raw.close()

        # 诊断：统计未匹配到网店订单号的行数
        unmatched_count = sum(1 for r in matched_rows if not r.get('web_order_id'))
        warn_msg = None
        if unmatched_count > 0:
            total_rows = len(matched_rows)
            warn_msg = f'注意：{unmatched_count}/{total_rows} 行未从sheet0匹配到网店订单号，已按订单编号单独分组。请检查sheet0是否包含对应的订单编号。'

        # 加载货品信息并核对
        products = load_products(PRODUCTS_JSON)
        barcode_to_info = build_barcode_to_info(products)
        code_to_info = build_code_to_info(products)
        results = verify_orders(matched_rows, barcode_to_info, code_to_info)

        # 导出结果
        download_id = f'result_{int(time.time())}.xlsx'
        output_path = os.path.join(OUTPUT_DIR, download_id)
        summary = export_results_to_xlsx(results, output_path)

        # 构建前端展示用的结果摘要
        order_results = []
        for r in results:
            has_warnings = bool(r['issues']) and r['is_correct']
            if r['is_correct'] and not has_warnings:
                status = 'correct'
            elif has_warnings:
                status = 'warning'
            else:
                status = 'error'

            order_results.append({
                'web_order_id': r['web_order_id'],
                'order_ids': r['order_ids'],
                'remark': r['remark'][:500] if r['remark'] else '',
                'status': status,
                'issues': r['issues'],
                'details': r['details'],
            })

        response_data = {
            'success': True,
            'summary': summary,
            'orders': order_results,
            'download_id': download_id,
        }
        if warn_msg:
            response_data['warning'] = warn_msg
        return jsonify(response_data)

    except Exception as e:
        import traceback
        traceback.print_exc()
        err_msg = str(e)
        # 对常见Excel列越界错误给出更友好的提示
        if 'tuple index out of range' in err_msg or 'index out of range' in err_msg:
            err_msg = 'Excel表格列数不足，请确保sheet0和sheet1的列结构与模板一致。常见原因：某行数据被删除或表格格式不正确。'
        return jsonify({'success': False, 'error': err_msg}), 500
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

@app.route('/api/download/<download_id>', methods=['GET'])
def download_result(download_id):
    """下载核对结果文件"""
    # 安全检查：只允许下载outputs目录下的文件
    if '..' in download_id or '/' in download_id or '\\' in download_id:
        return jsonify({'error': '非法文件名'}), 400
    path = os.path.join(OUTPUT_DIR, download_id)
    if not os.path.exists(path):
        return jsonify({'error': '文件不存在'}), 404
    return send_file(path, as_attachment=True, download_name=f'销售单核对结果_{download_id}')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
