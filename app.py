from flask import Flask, render_template, request, send_file, jsonify, flash
from werkzeug.utils import secure_filename
import os
import tempfile
import logging
from datetime import datetime
from main import markdown_to_docx_cli
from click.testing import CliRunner
import traceback

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # 在生产环境中应该使用环境变量

# 配置
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'md', 'txt'}
MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# 确保上传文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/.well-known/appspecific/com.chrome.devtools.json')
def chrome_devtools():
    """处理Chrome开发者工具的请求，避免404错误"""
    return jsonify({}), 200

@app.route('/convert', methods=['POST'])
def convert_markdown():
    try:
        # 检查是否有文件上传
        if 'file' not in request.files:
            return jsonify({'error': '没有选择文件'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '没有选择文件'}), 400
        
        if not allowed_file(file.filename):
            return jsonify({'error': '只支持 .md 和 .txt 文件'}), 400
        
        # 获取表单参数
        font_size = request.form.get('font_size', type=float)
        line_spacing = request.form.get('line_spacing', type=float)
        
        # 保存上传的文件
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        input_filename = f"{timestamp}_{filename}"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        file.save(input_path)
        
        # 创建输出文件路径
        output_filename = f"{timestamp}_{os.path.splitext(filename)[0]}.docx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        # 准备CLI参数
        cli_args = [input_path, output_path]
        if font_size:
            cli_args.extend(['--font-size', str(font_size)])
        if line_spacing:
            cli_args.extend(['--line-spacing', str(line_spacing)])
        
        # 使用CliRunner执行转换
        runner = CliRunner()
        result = runner.invoke(markdown_to_docx_cli, cli_args)
        
        if result.exit_code != 0:
            # 清理临时文件
            if os.path.exists(input_path):
                os.remove(input_path)
            return jsonify({'error': f'转换失败: {result.output}'}), 500
        
        # 检查输出文件是否存在
        if not os.path.exists(output_path):
            if os.path.exists(input_path):
                os.remove(input_path)
            return jsonify({'error': '转换失败: 输出文件未生成'}), 500
        
        # 清理输入文件
        if os.path.exists(input_path):
            os.remove(input_path)
        
        return jsonify({
            'success': True,
            'message': '转换成功！',
            'download_url': f'/download/{output_filename}'
        })
        
    except Exception as e:
        app.logger.error(f"转换过程中发生错误: {str(e)}")
        app.logger.error(traceback.format_exc())
        return jsonify({'error': f'服务器错误: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if not os.path.exists(file_path):
            return jsonify({'error': '文件不存在'}), 404
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        app.logger.error(f"下载文件时发生错误: {str(e)}")
        return jsonify({'error': '下载失败'}), 500

@app.route('/preview', methods=['POST'])
def preview_markdown():
    try:
        content = request.json.get('content', '')
        if not content.strip():
            return jsonify({'error': '内容不能为空'}), 400
        
        # 创建临时文件
        with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False, encoding='utf-8') as temp_file:
            temp_file.write(content)
            temp_input_path = temp_file.name
        
        # 创建输出文件路径
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"preview_{timestamp}.docx"
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
        
        # 获取参数
        font_size = request.json.get('font_size')
        line_spacing = request.json.get('line_spacing')
        
        # 准备CLI参数
        cli_args = [temp_input_path, output_path]
        if font_size:
            cli_args.extend(['--font-size', str(font_size)])
        if line_spacing:
            cli_args.extend(['--line-spacing', str(line_spacing)])
        
        # 执行转换
        runner = CliRunner()
        result = runner.invoke(markdown_to_docx_cli, cli_args)
        
        # 清理临时输入文件
        if os.path.exists(temp_input_path):
            os.remove(temp_input_path)
        
        if result.exit_code != 0:
            return jsonify({'error': f'转换失败: {result.output}'}), 500
        
        if not os.path.exists(output_path):
            return jsonify({'error': '转换失败: 输出文件未生成'}), 500
        
        return jsonify({
            'success': True,
            'message': '预览生成成功！',
            'download_url': f'/download/{output_filename}'
        })
        
    except Exception as e:
        app.logger.error(f"预览过程中发生错误: {str(e)}")
        return jsonify({'error': f'服务器错误: {str(e)}'}), 500

if __name__ == '__main__':
    # 设置日志
    logging.basicConfig(level=logging.INFO)
    # 生产环境关闭debug模式以提高性能
    app.run(debug=False, host='0.0.0.0', port=5000)
