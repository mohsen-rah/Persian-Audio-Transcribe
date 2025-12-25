<html lang="fa" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>تبدیل هوشمند صدا به متن فارسی - نسخه Ultra</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f6f8fa;
            color: #24292f;
            line-height: 1.6;
            margin: 0;
            padding: 20px;
        }
        .container {
            max-width: 900px;
            margin: 0 auto;
            background: white;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
            border: 1px solid #d0d7de;
        }
        h1, h2, h3 {
            color: #1f2328;
            border-bottom: 1px solid #d0d7de;
            padding-bottom: 10px;
        }
        h1 { font-size: 2em; text-align: center; border-bottom: none; }
        h2 { font-size: 1.5em; margin-top: 30px; }
        .header {
            text-align: center;
            margin-bottom: 40px;
        }
        .badges img {
            margin: 5px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        th, td {
            padding: 12px;
            border: 1px solid #d0d7de;
            text-align: right;
        }
        th {
            background-color: #f6f8fa;
        }
        code {
            background-color: #eff1f3;
            padding: 2px 6px;
            border-radius: 4px;
            font-family: Consolas, monospace;
            font-size: 0.9em;
            direction: ltr;
            display: inline-block;
        }
        pre {
            background-color: #f6f8fa;
            padding: 16px;
            border-radius: 6px;
            overflow-x: auto;
            direction: ltr;
            text-align: left;
        }
        .highlight {
            background-color: #e6ffec;
            border: 1px solid #48c774;
            color: #155724;
            padding: 15px;
            border-radius: 6px;
            margin: 20px 0;
        }
        details {
            background-color: #f6f8fa;
            border: 1px solid #d0d7de;
            border-radius: 6px;
            padding: 10px;
            margin-bottom: 10px;
            cursor: pointer;
        }
        summary {
            font-weight: bold;
            outline: none;
        }
        .footer {
            text-align: center;
            margin-top: 50px;
            font-size: 0.9em;
            color: #656d76;
        }
    </style>
</head>
<body>

<div class="container">
    <div class="header">
        <h1>🎙️ تبدیل هوشمند صدا به متن فارسی</h1>
        <h3>Persian Audio Transcriber - Ultra Edition</h3>
        
        <div class="badges">
            <img src="https://img.shields.io/badge/Python-3.x-blue?style=for-the-badge&logo=python&logoColor=white" alt="Python">
            <img src="https://img.shields.io/badge/FFmpeg-Included-green?style=for-the-badge&logo=ffmpeg&logoColor=white" alt="FFmpeg">
            <img src="https://img.shields.io/badge/License-MIT-red?style=for-the-badge" alt="License">
        </div>

        <p>
            <b>نسخه Ultra با قابلیت مانیتورینگ زنده، پردازش موازی و خروجی Word استاندارد</b><br>
            <i>تبدیل فایل‌های صوتی حجیم به متن فارسی با دقت بالا و مدیریت هوشمند خطا بدون نیاز به نصب پیچیده</i>
        </p>
    </div>

    <h2>✨ ویژگی‌های کلیدی</h2>
    <table>
        <thead>
            <tr>
                <th>ویژگی</th>
                <th>توضیحات</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>♾️ <b>بدون محدودیت</b></td>
                <td>قابلیت پردازش فایل‌های بسیار حجیم و طولانی (حتی چند ساعته) بدون قطع شدن.</td>
            </tr>
            <tr>
                <td>🎵 <b>فرمت‌های وسیع</b></td>
                <td>سازگاری کامل با <code>MP3</code>, <code>WAV</code>, <code>OGG</code>, <code>FLAC</code>, <code>M4A</code>, <code>AAC</code>, <code>WMA</code>.</td>
            </tr>
            <tr>
                <td>📊 <b>مانیتورینگ زنده</b></td>
                <td>دارای نوار پیشرفت (Progress Bar) دقیق با بروزرسانی هر ۰.۵ ثانیه.</td>
            </tr>
            <tr>
                <td>⚡ <b>سرعت و زمان</b></td>
                <td>نمایش لحظه‌ای سرعت پردازش (KB/s) و تخمین زمان اتمام (ETA).</td>
            </tr>
            <tr>
                <td>📝 <b>خروجی Word</b></td>
                <td>تولید فایل <code>.docx</code> با رعایت کامل راست‌چین (RTL) و فونت استاندارد.</td>
            </tr>
            <tr>
                <td>🛠 <b>اصلاح فارسی</b></td>
                <td>حل مشکل نمایش جدا جدا یا برعکس حروف فارسی در ترمینال ویندوز.</td>
            </tr>
        </tbody>
    </table>

    <h2>🚀 پیش‌نیازها (کامپیوتر)</h2>
    <p>برای اجرای صحیح برنامه روی ویندوز، مک یا لینوکس تنها به یک مورد نیاز دارید:</p>
    <div class="highlight">
        1️⃣ <b>Python 3.x</b>
    </div>
    <p><i>(نکته: فایل‌های اجرایی FFmpeg داخل پروژه قرار داده شده‌اند و <b>نیازی به دانلود یا نصب جداگانه ندارند</b>.)</i></p>

    <h2>📦 نصب و راه‌اندازی</h2>
    <p>۱. مخزن را کلون کنید یا فایل‌ها را دانلود نمایید.<br>
    ۲. با اجرای دستور زیر در ترمینال/CMD، تمام کتابخانه‌های مورد نیاز را نصب کنید:</p>
    <pre><code>pip install SpeechRecognition pydub python-docx colorama arabic-reshaper python-bidi</code></pre>

    <h2>🛠 نحوه استفاده</h2>
    
    <h3>روش اول: اجرای آسان (پیشنهادی برای ویندوز) ⚡</h3>
    <ol>
        <li>ابتدا فایل‌های صوتی خود را داخل پوشه <b><code>sot</code></b> کپی کنید.</li>
        <li>روی فایل <b><code>run.bat</code></b> دوبار کلیک کنید.</li>
        <li>تمام! برنامه اجرا شده و تبدیل را آغاز می‌کند.</li>
    </ol>

    <h3>روش دوم: از طریق ترمینال 💻</h3>
    <ol>
        <li>یک پوشه به نام <code>sot</code> در کنار فایل برنامه ایجاد کنید.</li>
        <li>فایل‌های صوتی را داخل آن بریزید.</li>
        <li>دستور زیر را اجرا کنید:</li>
    </ol>
    <pre><code>python transcribe.py</code></pre>

    <h2>⚙️ ساختار پروژه</h2>
    <pre><code>Project/
│
├── transcribe.py    # 🧠 موتور اصلی برنامه
├── run.bat          # ⚡ فایل اجرای آسان
├── ffmpeg.exe       # 🔧 موتور پردازش صدا
├── sot/             # 📂 ورودی فایل‌ها
└── outputs/         # 📄 خروجی Word</code></pre>

    <h2>❓ عیب‌یابی رایج</h2>
    
    <details>
        <summary>🔴 خطای FileNotFound یا WinError 2</summary>
        <p>این خطا زمانی رخ می‌دهد که فایل‌های <code>ffmpeg.exe</code> در کنار برنامه نباشند. مطمئن شوید این فایل‌ها در پوشه پروژه موجود هستند.</p>
    </details>

    <details>
        <summary>🔤 نمایش عجیب کلمات فارسی در CMD</summary>
        <p>اگر کلمات در محیط خط فرمان به صورت جدا جدا نمایش داده می‌شوند، دستور زیر را اجرا کنید:</p>
        <code>pip install arabic-reshaper python-bidi</code>
    </details>

    <div class="footer">
        <p>توسعه داده شده با ❤️<br>نسخه: Ultra 2.0</p>
    </div>
</div>

</body>
</html>
