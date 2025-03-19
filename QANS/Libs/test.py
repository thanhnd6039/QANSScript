import pyttsx3
from gtts import gTTS
class test(object):
    def convert(self):
        # Khởi tạo engine
        engine = pyttsx3.init()

        # Lấy danh sách các giọng nói có sẵn
        voices = engine.getProperty('voices')

        # Chọn giọng nam (thử tìm giọng nam có trong danh sách)
        for voice in voices:
            if "male" in voice.name.lower() or "nam" in voice.name.lower():
                engine.setProperty('voice', voice.id)
                break
        else:
            print("Không tìm thấy giọng nam, sử dụng giọng mặc định.")

        # Thiết lập tốc độ nói
        engine.setProperty('rate', 150)  # Mặc định là khoảng 200, giảm xuống cho giọng trầm hơn

        # Văn bản cần chuyển thành giọng nói
        text = "làm thế nào để phát triển chiều cao cho trẻ em"

        # Lưu thành file âm thanh
        engine.save_to_file(text, 'output.mp3')

        # Chạy engine để tạo file
        engine.runAndWait()

print("File âm thanh đã được tạo: output.mp3")

if __name__ == '__main__':
    var = test()
    run = var.convert()
