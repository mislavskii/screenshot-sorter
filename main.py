from screenshot_sorter import *


def main():
    x = Screens2PPTX(
        r'F:\User\Learn\ไทยศึกษา\เสียงสะท้อนอดีต\หลักฐานซากเรือ 1,200 ปี พบความเชื่อมโยงทางการค้าระหว่างรัฐทวารวดีกับศรีวิชัย.pptx',
        r'C:\Users\User\Videos\Captures'
    )
    x.build_title_slide()
    x.pull_images()


if __name__ == '__main__':
    main()
