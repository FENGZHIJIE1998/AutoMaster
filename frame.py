# 面板类
# 使用wxPython 教程 https://www.yiibai.com/wxpython
#
import datetime
import io

import wx
import os
import common
import pyWinhook
import win32api
import json
from run_script import RunScript


class Frame(wx.Frame):

    # 初始化
    def _init_frame(self, version, parent=None):
        wx.Frame.__init__(self,
                          id=wx.ID_ANY,
                          name='',
                          parent=parent,
                          pos=wx.Point(506, 283),
                          size=wx.Size(283, 205),
                          style=wx.STAY_ON_TOP | wx.DEFAULT_FRAME_STYLE,
                          title="Auto Master v%s" % version)
        self.SetClientSize(wx.Size(283, 205))
        self.SetBackgroundColour("#FFFFFF")
        self.HOT_KEYS = ['F9', 'F10', 'F11', 'F12']

        # 面板
        self.panel_main = wx.Panel(id=wx.ID_ANY, name='panel_main', parent=self,
                                   pos=wx.Point(0, 0), size=wx.Size(283, 205),
                                   style=wx.CAPTION)

        # 脚本Label
        self.script_label = wx.StaticText(id=wx.ID_ANY,
                                          label='脚本', name='script_label', parent=self.panel_main,
                                          pos=wx.Point(15, 15), size=wx.Size(40, 32), style=0)
        # 脚本下拉框
        self.script_choice = wx.Choice(choices=[], id=wx.ID_ANY,
                                       name='script_choice', parent=self.panel_main, pos=wx.Point(73, 10),
                                       size=wx.Size(121, 32), style=0)

        # 执行次数
        self.run_times = wx.SpinCtrl(id=wx.ID_ANY, initial=0, max=1000,
                                     min=0, name='run_times', parent=self.panel_main, pos=wx.Point(77, 48),
                                     size=wx.Size(47, 18), style=wx.SP_ARROW_KEYS)
        self.run_times.SetValue(1)

        # 执行次数提示语
        self.run_times_label = wx.StaticText(id=wx.ID_ANY,
                                             label='执行次数',
                                             name='run_times_label',
                                             parent=self.panel_main,
                                             pos=wx.Point(15, 49),
                                             size=wx.Size(50, 26), style=0)
        self.run_times_tips = wx.StaticText(id=wx.ID_ANY,
                                            label='0为无限循环',
                                            name='run_times_tips',
                                            parent=self.panel_main,
                                            pos=wx.Point(128, 49),
                                            size=wx.Size(74, 26), style=0)
        # 录制按键
        self.record_btn = wx.Button(id=wx.ID_ANY, label='录制',
                                    name='record_btn', parent=self.panel_main, pos=wx.Point(212, 7),
                                    size=wx.Size(56, 32), style=0)
        # 绑定事件
        self.record_btn.Bind(wx.EVT_BUTTON, self.OnRecordButton)

        # 启动按键
        self.run_btn = wx.Button(id=wx.ID_ANY, label='启动',
                                 name='run_btn', parent=self.panel_main, pos=wx.Point(212, 43),
                                 size=wx.Size(56, 32), style=0)
        # 绑定事件
        self.run_btn.Bind(wx.EVT_BUTTON, self.OnRunButton)

        self.start_label = wx.StaticText(id=wx.ID_ANY,
                                         label='启动热键', name='label_start_key',
                                         parent=self.panel_main, pos=wx.Point(16, 83), size=wx.Size(56, 24),
                                         style=0)

        self.stop_label = wx.StaticText(id=wx.ID_ANY,
                                        label='终止热键', name='label_stop_key',
                                        parent=self.panel_main, pos=wx.Point(16, 117), size=wx.Size(56, 32),
                                        style=0)

        # 启动热键
        self.start_choice = wx.Choice(choices=self.HOT_KEYS, id=wx.ID_ANY,
                                      name='choice_start', parent=self.panel_main, pos=wx.Point(73, 78),
                                      size=wx.Size(121, 32), style=0)
        self.start_choice.SetLabel('')
        self.start_choice.SetLabelText('')
        self.start_choice.Bind(wx.EVT_CHOICE, self.OnStartChoice,
                               id=wx.ID_ANY)
        # 停止热键
        self.stop_choice = wx.Choice(choices=self.HOT_KEYS, id=wx.ID_ANY,
                                     name='choice_stop', parent=self.panel_main, pos=wx.Point(73, 113),
                                     size=wx.Size(121, 32), style=0)
        self.stop_choice.Bind(wx.EVT_CHOICE, self.OnStopChoice,
                              id=wx.ID_ANY)

        # 状态提示语
        self.tips = wx.StaticText(id=wx.ID_ANY, label='请选择脚本',
                                  name='tips', parent=self.panel_main, pos=wx.Point(15, 155),
                                  size=wx.Size(100, 36), style=0)

        self.alert = wx.MessageDialog(self, message="Dialog", caption="警告")

    # 初始化面板
    def __init__(self, version, *args, **kw):
        super().__init__(*args, **kw)
        self._init_frame(version=version)
        self.SetIcon(common.get_icon())
        # 初始化脚本
        self.get_all_script()
        # 设置状态 默认0 运行：1 录制：2
        self.status = 0
        # 录制记录
        self.record = []
        # 操作时间Operation time
        self.o_time = common.current_ts()
        # 记录间隔时间，防止太多鼠标移动数据
        self.i_time = 200
        # 设置热键
        self.start_choice.SetSelection(0)
        self.stop_choice.SetSelection(1)
        # 创建管理对象
        self.hm = pyWinhook.HookManager()

        # 记录鼠标事件
        def on_mouse_event(event):

            # print('MessageName:',event.MessageName)  #事件名称
            # print('Message:',event.Message)          #windows消息常量
            # print('Time:',event.Time)                #事件发生的时间戳
            # print('Window:',event.Window)            #窗口句柄
            # print('WindowName:',event.WindowName)    #窗口标题
            # print('Position:',event.Position)        #事件发生时相对于整个屏幕的坐标
            # print('Wheel:',event.Wheel)              #鼠标滚轮
            # print('Injected:',event.Injected)        #判断这个事件是否由程序方式生成，而不是正常的人为触发。

            # 如果不在录制状态，直接结束
            if self.status != 2:
                return True

            # 获取事件名称
            message = event.MessageName
            # 仅监听以下事件
            all_messages = ('mouse left down', 'mouse left up', 'mouse right down', 'mouse right up', 'mouse move')
            if message not in all_messages:
                return True

            # 获取当前鼠标位置
            pos = win32api.GetCursorPos()
            # 两次调用的间隔时间
            delay = common.current_ts() - self.o_time

            # 如果只是鼠标移动，并且精度大于间隔时间，则不记录，直接结束
            if message == 'mouse move' and delay < self.i_time:
                return True
            # 如果记录，刷新记录时间
            self.o_time = common.current_ts()
            # 如果没有记录，代表首次记录，操作时间记为0
            if not self.record:
                delay = 0
            # 打印 操作时间，操作时间，操作位置
            print(delay, message, pos)
            # 追加脚本记录
            self.record.append([delay, 'EM', message, pos])

            # 这一块主要是在GUI上显示记录次数，非重点
            # 获取标签
            text = self.tips.GetLabel()
            # 设置记录？
            action_count = text.replace(' 个操作已记录', '')
            # 修改记录次数
            text = '%d 个操作已记录' % (int(action_count) + 1)
            # 重新设置文本
            self.tips.SetLabel(text)
            return True

        # 记录键盘事件
        def on_keyboard_event(event):

            # print('MessageName:', event.MessageName)  # 同上，共同属性不再赘述
            # print('Message:',event.Message)
            # print('Time:',event.Time)
            # print('Window:',event.Window)
            # print('WindowName:',event.WindowName)
            # print('Ascii:', event.Ascii, chr(event.Ascii))   #按键的ASCII码
            # print('Key:', event.Key)                         #按键的名称
            # print('KeyID:', event.KeyID)                     #按键的虚拟键值
            # print('ScanCode:', event.ScanCode)               #按键扫描码
            # print('Extended:', event.Extended)               #判断是否为增强键盘的扩展键
            # print('Injected:', event.Injected)
            # print('Alt', event.Alt)                          #是某同时按下Alt
            # print('Transition', event.Transition)            #判断转换状态

            # 获取事件名称
            message = event.MessageName
            # 简化事件名称
            message = message.replace(' sys ', ' ')
            # 热键            # 热键键键键
            if message == 'key up' and (self.status == 1 or self.status == 0):
                # 监听启动和停止事件
                # 获取按键
                key_name = event.Key.lower()
                # 获取按键名称
                start_name = self.HOT_KEYS[self.start_choice.GetCurrentSelection()].lower()
                stop_name = self.HOT_KEYS[self.stop_choice.GetCurrentSelection()].lower()

                if key_name == start_name and self.status == 0:
                    print('热键启动')
                    t = RunScript(self)
                    t.start()

                    # 如果按键是停止键，同时在运行状态，停止代码
                elif key_name == stop_name and self.status == 1:
                    # 如果正在执行脚本，并且按下了F9，终止进程
                    print('热键停止')
                    self.status = 0

            # 如果不是在录制状态，结束
            if self.status != 2:
                return True
            # 监听的事件
            all_messages = ('key down', 'key up')
            # 如果不是想要监听的事件，结束
            if message not in all_messages:
                return True
            # 按键的数据
            key_info = (event.KeyID, event.Key, event.Extended)

            # 两次调用的间隔时间
            delay = common.current_ts() - self.o_time

            self.o_time = common.current_ts()
            # 第一次，记0
            if not self.record:
                delay = 0
            # 打印脚本
            print(delay, message, key_info)
            # 录制操作
            self.record.append([delay, 'EK', message, key_info])

            # 这一块主要是在GUI上显示记录次数，非重点
            # 获取标签
            text = self.tips.GetLabel()
            # 设置记录？
            action_count = text.replace(' 个操作已记录', '')
            # 修改记录次数
            text = '%d 个操作已记录' % (int(action_count) + 1)
            # 重新设置文本
            self.tips.SetLabel(text)
            return True

        self.hm.MouseAll = on_mouse_event
        self.hm.KeyAll = on_keyboard_event
        self.hm.HookMouse()
        self.hm.HookKeyboard()

    # 获取脚本路径
    def get_script_path(self):
        # 从下拉列表获取脚本index
        i = self.script_choice.GetSelection()
        # 无脚本为空
        if i < 0:
            return ''
        # 获取选中的脚本
        script = self.scripts[i]
        # 拼接路径
        path = os.path.join(os.getcwd(), 'scripts', script)
        print(path)
        return path

    # 创建新脚本
    def new_script_path(self):
        now = datetime.datetime.now()
        # 设置脚本文件名
        script = '%s.txt' % now.strftime('%m%d_%H%M')
        # 重名另起
        if script in self.scripts:
            script = '%s.txt' % now.strftime('%m%d_%H%M%S')
        # 把新文件追加到下拉列表第一个
        self.scripts.insert(0, script)
        # 并选中它
        self.script_choice.SetItems(self.scripts)
        self.script_choice.SetSelection(0)
        # 返回路径
        return self.get_script_path()

    # 获取当前路径下的脚本文件夹里的所有脚本
    def get_all_script(self):
        # 如果不存在，则创建
        if not os.path.exists('scripts'):
            os.mkdir('scripts')
        self.scripts = os.listdir('scripts')[::-1]
        # 筛选出txt文件
        self.scripts = list(filter(lambda s: s.endswith('.txt'), self.scripts))
        self.script_choice.SetItems(self.scripts)
        # 如果存在脚本文件，默认选择第一个
        if self.scripts:
            self.script_choice.SetSelection(0)

    # 录制按钮点击事件
    def OnRecordButton(self, event):
        # 如果在录制中中，停止记录
        if self.status == 2:
            print('录制停止..')
            self.status = 0
            # 删掉最后两条停止录制的操作
            self.record = self.record[:-2]
            # 获取输出格式
            output = json.dumps(self.record, indent=1)
            # 格式化
            output = output.replace('\r\n', '\n').replace('\r', '\n')
            output = output.replace('\n   ', '').replace('\n  ', '')
            output = output.replace('\n ]', ']')
            # 写入文件
            open(self.new_script_path(), 'w').write(output)
            # 修改录制按键的文字
            self.record_btn.SetLabel('录制')
            self.tips.SetLabel('录制已结束')
            # 清空记录
            self.record = []
        elif self.status == 0:
            # 开始记录
            print('录制开始')
            self.status = 2
            self.o_time = common.current_ts()

            # 修改录制按键的文字
            self.record_btn.SetLabel('结束')  # 结束
            self.tips.SetLabel('0 个操作已记录')
            # 将下拉变成空白
            self.script_choice.SetSelection(-1)
            self.record = []
        else:
            self.tips.SetLabel('脚本执行中，不允许录制')
        # 不处理事件让后续处理
        event.Skip()

    # 启动按钮点击事件
    def OnRunButton(self, event):
        t = RunScript(self)
        t.start()
        event.Skip()

    def OnStartChoice(self, event):
        if self.start_choice.GetCurrentSelection() == self.stop_choice.GetCurrentSelection():
            self.start_choice.SetSelection(0)
            self.alert.Message = "热键冲突，请重新设置"
            self.alert.ShowModal()
        if self.start_choice.GetCurrentSelection() == self.stop_choice.GetCurrentSelection():
            self.start_choice.SetSelection(1)
        event.Skip()

    def OnStopChoice(self, event):
        if self.start_choice.GetCurrentSelection() == self.stop_choice.GetCurrentSelection():
            self.stop_choice.SetSelection(1)
            self.alert.Message = "热键冲突，请重新设置"
            self.alert.ShowModal()
        if self.start_choice.GetCurrentSelection() == self.stop_choice.GetCurrentSelection():
            self.stop_choice.SetSelection(0)
        event.Skip()


if __name__ == '__main__':
    # # HELLO WORLD
    # app = wx.App()
    # window = wx.Frame(None, title = "wxPython - www.yiibai.com", size = (400,300))
    # panel = wx.Panel(window)
    # label = wx.StaticText(panel, label = "Hello World", pos = (100,100))
    # window.Show(True)
    # app.MainLoop()
    app = wx.App()
    frm = Frame("1.0")
    frm.Show()
    app.MainLoop()
