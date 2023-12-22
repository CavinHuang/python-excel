import asyncio
from pyppeteer import launch
from lxml import etree
import requests
import ddddocr
from anti_useragent import UserAgent
import os


desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
local_image_tmp = os.path.join(desktop_path, '_tmp-images')


# 识别验证码
def find_image(code_path):
    ocr = ddddocr.DdddOcr()
    with open(code_path, 'rb') as f:
        img_bytes = f.read()
    result = ocr.classification(img_bytes)
    return result
 
ua = UserAgent() 
async def main(url, file_full_name, logger):
    '''
    :param ASIN: 商品的对应asin
    :return:
    '''
    url = f'{url}'
    browser = await launch({
        "headless": True,
        "defaultViewport": None,
        "args": ['--lang=en-US'],
        "handleSIGINT": False,
            "handleSIGTERM": False,
            "handleSIGHUP": False
    })
    pages = await browser.pages()
    page = pages[0]
    page.setUserAgent(ua.random)
    await page.goto(url, { 'timeout': 0 })
 
    while True:
        await asyncio.sleep(2)
        htm = await page.content()
        html = etree.HTML(htm)
        image = html.xpath('//div[@class="a-row a-text-center"]/img/@src')
        if image:
            logger.info('出现验证码')
            # 获取图片二进制数据
            headers = {'Users-Agent': ua.random}
            img_data = requests.get(url=image[0], headers=headers).content
            code_path = os.path.join(local_image_tmp, 'code.png')
            with open(code_path, 'wb') as fp:
                fp.write(img_data)
 
            # 在输入框输入验证码
            # await asyncio.sleep(2)
            result = find_image(code_path)
            await page.hover('input#captchacharacters')
            await page.type('input#captchacharacters', result)
            # 点击确定按钮
            # await asyncio.sleep(2)
            await page.hover('button[type="submit"]')
            await page.click('button[type="submit"]')
 
        else:
            logger.info('没有出现验证码, 进入页面成功')
            break
 
    await asyncio.sleep(2)

    # 获取网页信息
    htm = await page.content()
    html = etree.HTML(htm)
    image = html.xpath('//*[@id="landingImage"]/@src')
    
    if image:
        logger.info('获取图片')
        # 获取图片二进制数据
        headers = {'Users-Agent': ua.random}
        img_data = requests.get(url=image[0], headers=headers).content
        with open(file_full_name, 'wb') as fp:
            fp.write(img_data)
    else:
        print('页面图片查找失败')
        with open('./example.html', 'wb') as fp:
            fp.write(etree.tostring(html, encoding='utf-8', pretty_print=True))
    await browser.close()
 
# 设置一个商品asin
#ASIN = 'https://www.amazon.co.uk/dp/B0BYRKFH4J?th=1'
#asyncio.get_event_loop().run_until_complete(main(ASIN))

def fetch_img(url, row_index, logger):
    file_name = f"No{row_index}_row_image.png"
    file_full_name = os.path.join(local_image_tmp, file_name)
    if not os.path.exists(local_image_tmp):
        # 如果不存在，递归地创建文件夹
        os.makedirs(local_image_tmp)
        print(f"文件夹 {local_image_tmp} 创建成功")
    else:
        print(f"文件夹 {local_image_tmp} 已存在")

    _loop = asyncio.new_event_loop()
    asyncio.set_event_loop(_loop)
    _loop.run_until_complete(main(url, file_full_name, logger))
    print(row_index, url, file_full_name)
    _loop.close()
    return file_full_name