import configparser
from google import genai

def test_gemini_api():
    """测试Gemini API连接"""
    try:
        # 加载配置
        config = configparser.ConfigParser()
        config.read('config.ini', encoding='utf-8')
        api_key = config.get('DEFAULT', 'api_key')
        
        # 初始化客户端
        client = genai.Client(api_key=api_key)
        print("✓ Gemini API客户端初始化成功")
        
        # 测试简单翻译
        response = client.models.generate_content(
            model="gemini-2.0-flash",
            contents="将这句话翻译为英文：你好世界"
        )
        
        print(f"✓ API测试成功")
        print(f"测试翻译结果: {response.text}")
        
        return True
        
    except Exception as e:
        print(f"✗ API测试失败: {e}")
        return False

if __name__ == "__main__":
    print("正在测试Gemini API连接...")
    test_gemini_api() 