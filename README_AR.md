# إصلاح Khorfakkan Portal — إخفاء Google Apps Script من الواجهة

هذه النسخة تعالج السبب الحقيقي للمشكلة الحالية.

## لماذا كان يظهر `script.google.com`؟

النسخة الحالية في GitHub تستخدم iframe يشير مباشرة إلى:

https://script.google.com/macros/s/AKfycbxH3Q99G0dM0xxWwKSOO3Lds9HJnPi0UG_JCpzHXRAEx3o2tf18mI0Zfws7T0bjsKlNLg/exec

ولذلك المستخدم يدخل فعليًا إلى Google Apps Script أو يرى واجهته، وتظهر رسالة:

"This application was created by a Google Apps Script user"

الحل هنا ليس Redirect ولا iframe.

## ما الذي تغير؟

- Cloudflare Pages يستقبل `/dashboard`, `/requests`, `/admin`, `/statistics`, `/assistant`.
- Cloudflare يعمل Reverse Proxy إلى Google Apps Script.
- `google.script.run` تم استبداله بواجهة RPC داخلية:
  `browser -> /api/rpc -> Google Apps Script`
- روابط Apps Script داخل الصفحات يعاد كتابتها إلى دومين Cloudflare.
- صفحة `/voice` تعمل مباشرة من Cloudflare Pages.
- إذا كان رابط التسجيل الصوتي غير مضبوط، يتم استخدام `/voice` تلقائيًا.
- تم إضافة `doPost(e)` إلى `Code.gs` لاستقبال RPC.

## الملفات

ضع في جذر GitHub:

- `index.html`
- `voice.html`
- `functions/[[path]].js`

وفي Google Apps Script استبدل `Code.gs` بالنسخة الموجودة هنا.

## خطوات النشر

### 1) Google Apps Script

استبدل `Code.gs` الحالي بالملف الموجود في هذه الحزمة.

ثم:

Deploy
→ Manage deployments
→ Edit
→ New version
→ Deploy

ويجب أن يكون Web App متاحًا للمستخدمين الذين يستخدمون النظام.

### 2) GitHub

في:

`sfaroukf/khorfakkan-portal`

احذف `index.html` القديم واستبدله بالنسخة الموجودة هنا.

أضف:

`functions/[[path]].js`

واستبدل `voice.html` بالنسخة الموجودة هنا.

ثم Commit & push.

Cloudflare Pages مربوط بالفعل بالمستودع، لذلك سيبدأ Deployment تلقائيًا.

### 3) الاختبار

استخدم دومين Cloudflare فقط:

https://khorfakkan-portal.pages.dev/

ثم:

https://khorfakkan-portal.pages.dev/dashboard

https://khorfakkan-portal.pages.dev/requests

https://khorfakkan-portal.pages.dev/admin

https://khorfakkan-portal.pages.dev/statistics

https://khorfakkan-portal.pages.dev/assistant

https://khorfakkan-portal.pages.dev/voice

## مهم جدًا

لا تختبر باستخدام رابط:

https://script.google.com/macros/s/AKfycbxH3Q99G0dM0xxWwKSOO3Lds9HJnPi0UG_JCpzHXRAEx3o2tf18mI0Zfws7T0bjsKlNLg/exec

لأن هذا هو رابط Google Apps Script نفسه، وسيظل يعرض رسالة Google مهما فعلنا في Cloudflare.

بعد تركيب النسخة، يجب أن يكون شريط العنوان دائمًا على:

`khorfakkan-portal.pages.dev`

ولا يجب أن ينتقل إلى `script.google.com`.

## ملاحظة تسجيل الدخول

النظام الحالي يدعم تسجيل الدخول بحساب الموظف ورمز `staffToken`. هذا المسار يعمل مع الـ Proxy.

جلسة Google الشخصية (`Session.getActiveUser()`) لا تنتقل إلى Cloudflare، لأن Cloudflare ليس نطاق Google. لذلك إذا كنت تعتمد على الدخول التلقائي بحساب Google المالك، استخدم تسجيل دخول الموظف داخل النظام أو ننقل المصادقة لاحقًا إلى Cloudflare/Supabase.
