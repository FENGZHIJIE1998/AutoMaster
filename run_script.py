# 执行脚本
import json
import threading
import time
import traceback

import win32api
import win32con


class RunScript(threading.Thread):

    # 初始化
    def __init__(self, frame):
        # 设置窗口
        self.frame = frame
        super(RunScript, self).__init__()

    # 启动
    def run(self):
        # 如果运行或者录制，结束
        if self.frame.status != 0:
            return
        # 获取脚本路径
        script_path = self.frame.get_script_path()
        # 如果无脚本文件，提示先录制一个
        if not script_path:
            self.frame.tips.SetLabel('你还没有可用脚本，请先录制')
            return
        # 开始运行
        self.frame.status = 1
        try:
            # 获取运行次数
            run_times = self.frame.run_times.Value
            running_text = '%s 正在执行.. ' % script_path.split('/')[-1].split('\\')[-1]
            self.frame.tips.SetLabel(running_text)

            script = self.get_script(script_path)
            j = 0
            # 循环运行，如果为0，无限循环
            while j < run_times or run_times == 0:
                j += 1
                if self.frame.status != 1:
                    break
                self.run_script_once(script)
            # 修改状态
            if self.frame.status == 1:
                # =1意味着正常运行结束，否则意味着强制结束
                self.frame.tips.SetLabel('脚本已运行完毕')
            else:
                self.frame.tips.SetLabel('脚本已强制终止')
            self.frame.status = 0
            print('脚本已运行完毕')

        except Exception as e:
            # 抛出异常
            print('出错啦！', e)
            traceback.print_exc()
            self.frame.tips.SetLabel('出错啦！')
            self.frame.status = 0

    # 获取脚本数据
    def get_script(self, script_path):
        content = ''
        lines = []

        # 读取脚本
        try:
            lines = open(script_path, 'r', encoding='utf8').readlines()
        except Exception as e:
            print(e)
            try:
                lines = open(script_path, 'r', encoding='gbk').readlines()
            except Exception as e:
                print(e)

        for line in lines:
            # 去注释
            if '//' in line:
                index = line.find('//')
                line = line[:index]
            # 去空字符
            line = line.strip()
            content += line

        # 去最后一个元素的逗号（如有）
        content = content.replace('],\n]', ']\n]').replace('],]', ']]')

        print(content)
        # 加载数据
        return json.loads(content)

    def run_script_once(self, script):

        # 加载数据
        steps = len(script)  # 获取操作数量
        # 获取当前电脑分辨率，更好适配
        sw = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
        sh = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
        # 依次执行脚本
        for i in range(steps):
            if self.frame.status != 1:
                break
            print(script[i])
            delay = script[i][0]
            event_type = script[i][1].upper()
            message = script[i][2].lower()
            action = script[i][3]
            # 控制执行速度
            time.sleep(delay / 1000.0)

            # 如果事件类型为鼠标
            if event_type == 'EM':
                x, y = action
                # 保持不动、跳过
                if action == [-1, -1]:
                    # 约定 [-1, -1] 表示鼠标保持原位置不动
                    pass
                else:
                    # 挪动鼠标
                    nx = int(x * 65535 / sw)
                    ny = int(y * 65535 / sh)
                    win32api.mouse_event(win32con.MOUSEEVENTF_ABSOLUTE | win32con.MOUSEEVENTF_MOVE, nx, ny, 0, 0)
                # 点击事件
                if message == 'mouse left down':
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                elif message == 'mouse left up':
                    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                elif message == 'mouse right down':
                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
                elif message == 'mouse right up':
                    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                elif message == 'mouse move':
                    pass
                else:
                    print('非法鼠标事件:', message)
            elif event_type == "EK":
                key_code, key_name, extended = action

                # shift ctrl alt
                # if key_code >= 160 and key_code <= 165:
                #     key_code = int(key_code/2) int(key_code/2) int(key_code/2) - 64

                base = 0
                if extended:
                    base = win32con.KEYEVENTF_EXTENDEDKEY

                if message == 'key down':
                    win32api.keybd_event(key_code, 0, base, 0)
                elif message == 'key up':
                    win32api.keybd_event(key_code, 0, base | win32con.KEYEVENTF_KEYUP, 0)
                else:
                    print('非法键盘事件:', message)
