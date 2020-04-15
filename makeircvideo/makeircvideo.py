import os
import sys

povray_path = 'Your povray PATH'
ffmpeg_path = 'Your ffmpeg PATH'


def get_resolution():
    for filename in sorted(os.listdir()):
        if filename.endswith('.pov'):
            with open(filename, errors=None, mode='r') as file:
                for index, line in enumerate(file, 1):
                    if index == 2:
                        width = line.split()[2]
                    if index == 3:
                        height = line.split()[2]
                        break
    return [width, height]


def do_povray(width, height):
    for filename in sorted(os.listdir()):
        if filename.endswith('.pov'):
            command = '%s +W%s +H%s +A %s' % (povray_path,
                                              str(width), str(height), filename)
            os.system(command)


def do_ffmpeg(filename, switch):
    if switch == 0:  # 渲染mp4视频
        os.system('%s -r 15 -i FRAME%s.png -crf 22 %s.mp4' %
                  (ffmpeg_path, '%04d', filename))
    elif switch == 1:  # gif
        os.system('%s -i FRAME0001.png -vf palettegen palette.png' %
                  (ffmpeg_path))
        os.system('%s -r 15 -i FRAME%s.png -i palette.png -lavfi paletteuse %s.gif' %
                  (ffmpeg_path, '%04d', filename))
    else:
        print("Wrong switch used!")


def main():
    filename = os.getcwd().split('\\')[-1][:-8]
    switch = 0
    if len(sys.argv) != 1:
        if sys.argv[1] == '1' or sys.argv[1] == '2':
            switch = int(sys.argv[1]) - 1
    [width, height] = get_resolution()
    if len(sys.argv) == 3 and sys.argv[2] == '1':
        pass
    else:
        do_povray(width, height)
    do_ffmpeg(filename, switch)
    print("Program executed successfully!\n")
    exit(0)


if __name__ == "__main__":
    print("Initializing...")
    main()
