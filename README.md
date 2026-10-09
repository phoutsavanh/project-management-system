# ລະບົບບໍລິຫານຈັດການໂຄງການ (Project Management System)

ລະບົບບໍລິຫານຈັດການໂຄງການ, ກິດຈະກຳ, ງົບປະມານ ແລະ ການເບີກຈ່າຍ ທີ່ໃຊ້ Google Sheets ເປັນຖານຂໍ້ມູນ ແລະ ສາມາດ Deploy ຜ່ານ **Vercel** ຫຼື **Google Apps Script** ໄດ້ໂດຍກົງ.

---

## 🚀 ວິທີການ Deploy ໃນ Vercel

### ຂັ້ນຕອນທີ 1: Deploy Google Apps Script (Backend)
1. ເຂົ້າໄປທີ່ Google Apps Script ຂອງ Sheet ທ່ານ.
2. ນຳເອົາໂຄດໃນໄຟລ໌ `code.gs` ໄປວາງທັບໃນ Apps Script.
3. ກົດ **Deploy** > **Manage deployments** > ກົດໄອຄອນສໍ ✏️ (Edit) > ເລືອກ **Version: New version**.
4. ຕັ້ງຄ່າ:
   - **Execute as:** `Me (ອີເມວຂອງທ່ານ)`
   - **Who has access:** `Anyone (ທຸກຄົນ)`
5. ກົດ **Deploy** ແລ້ວ **Copy Web App URL** (ລິ້ງທີ່ລົງທ້າຍດ້ວຍ `/exec`).

### ຂັ້ນຕອນທີ 2: ເຊື່ອມຕໍ່ URL ກັບ Frontend
ທ່ານສາມາດເຮັດໄດ້ 2 ວິທີ:
- **ວິທີທີ 1:** ເປີດໄຟລ໌ `config.js` ແລ້ວວາງ URL ໃສ່ໃນ:
  ```javascript
  window.DEFAULT_GAS_API_URL = "https://script.google.com/macros/s/.../exec";
  ```
- **ວິທີທີ 2:** ບໍ່ຕ້ອງແກ້ໄຂໂຄດ, ເມື່ອເປີດໜ້າເວັບໃນ Vercel ໃຫ້ກົດປຸ່ມ **⚙️ ຕັ້ງຄ່າ API URL** ໃນໜ້າ Login ແລ້ວວາງ URL ໃສ່.

### ຂັ້ນຕອນທີ 3: Push ຂຶ້ນ GitHub
```bash
git init
git add .
git commit -m "Initial commit for Vercel deployment"
git branch -M main
git remote add origin https://github.com/<YOUR_USERNAME>/<YOUR_REPOSITORY>.git
git push -u origin main
```

### ຂັ້ນຕອນທີ 4: Deploy ໃນ Vercel
1. ເຂົ້າໄປທີ່ [vercel.com](https://vercel.com) ແລ້ວ Login ດ້ວຍ GitHub.
2. ກົດ **Add New...** > **Project**.
3. ເລືອກ Repository ທີ່ທ່ານຫາກໍ push ຂຶ້ນໄປ.
4. ກົດ **Deploy** (ບໍ່ຕ້ອງປ່ຽນຄ່າໃດໆ ເພາະລະບົບເປັນ HTML/JS ແບບ Static).
5. ລໍຖ້າ 10-20 ວິນາທີ ຈະໄດ້ລິ້ງ Domain ຂອງ Vercel (ເຊັ່ນ `https://your-project.vercel.app`) ພ້ອມໃຊ້ງານທັນທີ!

---

## 🔐 ຂໍ້ມູນເຂົ້າສູ່ລະບົບເລີ່ມຕົ້ນ
- **Username:** `admin`
- **Password:** `1234`
