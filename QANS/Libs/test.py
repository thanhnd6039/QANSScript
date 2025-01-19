from gtts import gTTS

class test(object):
    def convert(self):
        text = ("Đây là bốn bí quyết giúp trẻ phát triển chiều cao hiệu quả. Thứ nhất là chế độ dinh dưỡng khoa học. "
                "Thứ hai là luyện tập thể dục thể thao. "
                "Thứ ba là có một giấc ngủ chất lượng. "
                "Và cuối cùng là bổ sung Mi đu Mena Q bảy 180.")
        language = "vi"
        tts = gTTS(text=text, lang=language, slow=False)
        output_file = "output.mp3"
        tts.save(output_file)
        print(f"File âm thanh đã được lưu thành công: {output_file}")


if __name__ == '__main__':
    var = test()
    run = var.convert()
