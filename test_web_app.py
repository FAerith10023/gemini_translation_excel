#!/usr/bin/env python
"""
Web应用功能测试脚本
"""

import requests
import os
import time

def test_web_app():
    """测试Web应用的各项功能"""
    base_url = "http://localhost:5000"
    
    print("🧪 开始测试Web应用功能...")
    print("=" * 50)
    
    # 测试1: 首页访问
    print("1️⃣ 测试首页访问...")
    try:
        response = requests.get(base_url)
        if response.status_code == 200:
            print("   ✅ 首页访问成功")
        else:
            print(f"   ❌ 首页访问失败，状态码: {response.status_code}")
    except Exception as e:
        print(f"   ❌ 连接失败: {e}")
        return False
    
    # 测试2: API连接测试
    print("\n2️⃣ 测试API连接...")
    try:
        response = requests.get(f"{base_url}/test_api")
        data = response.json()
        if data.get('success'):
            print("   ✅ API连接正常")
            print(f"   📝 响应: {data.get('response', '')[:50]}...")
        else:
            print(f"   ❌ API连接失败: {data.get('error')}")
    except Exception as e:
        print(f"   ❌ API测试失败: {e}")
    
    # 测试3: 文件上传测试
    print("\n3️⃣ 测试文件上传...")
    if os.path.exists('input.xlsx'):
        try:
            with open('input.xlsx', 'rb') as f:
                files = {'file': ('test.xlsx', f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
                response = requests.post(f"{base_url}/upload", files=files)
                
            data = response.json()
            if data.get('success'):
                print("   ✅ 文件上传成功")
                print(f"   📄 文件名: {data.get('filename')}")
                print(f"   📊 行数: {data.get('max_row')}, 列数: {data.get('max_col')}")
                print(f"   🆔 文件ID: {data.get('file_id')}")
                return data.get('file_id')
            else:
                print(f"   ❌ 文件上传失败: {data.get('error')}")
        except Exception as e:
            print(f"   ❌ 上传测试失败: {e}")
    else:
        print("   ⚠️ 测试文件input.xlsx不存在，跳过上传测试")
    
    return None

def test_translation(file_id):
    """测试翻译功能"""
    if not file_id:
        print("\n4️⃣ 跳过翻译测试（无文件ID）")
        return
    
    print("\n4️⃣ 测试翻译功能...")
    base_url = "http://localhost:5000"
    
    translate_data = {
        'file_id': file_id,
        'range': 'A1:C5',
        'source_lang': '中文',
        'target_lang': '英文',
        'custom_prompt': '量词"个"翻译为"nos"'
    }
    
    try:
        response = requests.post(
            f"{base_url}/translate",
            json=translate_data,
            headers={'Content-Type': 'application/json'}
        )
        
        data = response.json()
        if data.get('success'):
            print("   ✅ 翻译功能正常")
            print(f"   📊 翻译统计: {data.get('translated_count')}/{data.get('total_count')}")
            print(f"   💾 输出文件: {data.get('output_file')}")
            print(f"   📥 下载链接: {data.get('download_url')}")
        else:
            print(f"   ❌ 翻译失败: {data.get('error')}")
    except Exception as e:
        print(f"   ❌ 翻译测试失败: {e}")

def main():
    """主测试函数"""
    print("🌐 Excel翻译工具 Web版 - 功能测试")
    print("请确保Web服务器已启动（python run_web.py）")
    print()
    
    # 等待用户确认
    input("按回车键开始测试...")
    
    # 执行测试
    file_id = test_web_app()
    
    if file_id:
        # 询问是否测试翻译功能
        test_translate = input("\n是否测试翻译功能？(y/n): ").strip().lower()
        if test_translate == 'y':
            test_translation(file_id)
    
    print("\n" + "=" * 50)
    print("🎉 测试完成！")
    print("\n🔗 Web界面地址: http://localhost:5000")
    print("📖 使用说明请查看 README.md")

if __name__ == '__main__':
    main() 