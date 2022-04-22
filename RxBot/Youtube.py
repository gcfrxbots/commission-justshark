from pywinauto import Application, ElementAmbiguousError


def detectUrl():
    try:
        app = Application(backend='uia')
        app.connect(title_re=".*Chrome.*")
        dlg = app.top_window()
        url = dlg.child_window(title="Address and search bar", control_type="Edit").get_value()
        return url
    except ElementAmbiguousError:
        print("You have multiple chrome browsers open, the bot can only detect the URL if you only have one window open!")
    except:
        pass


def detectedNewVideo(url):
    print(url)