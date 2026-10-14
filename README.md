# Inventory Manager (Python + Excel)

Stock IN / OUT, profit report, aur unpaid hisaab. Saara data `inventory.xlsx` me save hota hai.

## Pehli baar setup (VS Code)
1. Python 3.9 ya naya install karo (python.org). Install karte waqt "Add Python to PATH" tick karo.
2. VS Code me **File > Open Folder** se `inventory-app` folder kholo.
3. Terminal kholo (**Terminal > New Terminal**) aur chalao:
   ```
   pip install -r requirements.txt
   ```

## Chalana
```
python app.py
```
Phir browser me kholo: **http://localhost:5000**
(Windows par `start.bat` par double click bhi kar sakte ho.)

## Phone se use karna (camera scan ke saath)
1. Phone aur computer **same Wi-Fi** par hon.
2. Camera ke liye https chahiye, to aise chalao:
   ```
   python app.py --https
   ```
3. Terminal me jo "On your phone" wala address dikhe (jaise `https://192.168.1.5:5000`) wo phone ke browser me kholo.
4. "Not secure" ka warning aaye to **Advanced > Continue** dabao. Ye apna hi computer hai, safe hai.
5. Na khule to Windows Firewall me Python ko allow karo.

Bina camera wala "Without camera" mode aur USB scanner http par bhi chalte hain.

## Data kahan hai
- `inventory.xlsx` me 3 sheets hain: **Items** (stock aur prices), **History** (har IN/OUT) aur **Unpaid** (baaki paisa).
- `backups/` folder me roz ki pehli save par us din ki copy ban jati hai.
- Excel file ko **band karke** app use karo. Windows me Excel me khuli file par app save nahi kar pata, tab "NOT SAVED" ka message aata hai.
- Excel me Items sheet ki values (stock, price) badal sakte ho, par pehli row (heading) aur sheet ke naam mat badlo. "Unpaid" sheet sirf dekhne ke liye hai, app use dobara bana deta hai.
- Naya shuru karna ho to app me History tab > "Reset all data" dabao (backup apne aap ban jata hai).

## Dhyan rakhne wali baat
- Ek time par ek hi device se entry karo. Do phone se saath me entry karoge to ek ki save hui entry dusre se overwrite ho sakti hai. Page dobara kholne par latest data aa jata hai.
- Computer band hoga to app bhi band rahega. Phone se use karne ke liye computer chalu aur app running chahiye.
- Camera scan ke liye internet chahiye (scanner library online se aati hai). Type/search mode bina internet ke chalta hai.

## Backup ko cloud se jodna (Google Drive / OneDrive)
1. Google Drive for desktop (ya OneDrive) install karo aur usme ek folder banao, jaise `InventoryBackup`.
2. `start.bat` kholo, `rem set INVENTORY_BACKUP_DIR=...` wali line se `rem` hatao aur apne folder ka path likho.
   Terminal se chalate ho to:
   - PowerShell: `$env:INVENTORY_BACKUP_DIR="G:\My Drive\InventoryBackup"; python app.py`
   - CMD: `set INVENTORY_BACKUP_DIR=G:\My Drive\InventoryBackup` phir `python app.py`
3. Ab roz ki backup copy usi folder me banegi aur Drive use apne aap cloud par bhej dega.
4. Mukhya `inventory.xlsx` ko sync folder me mat rakho (sync ke time file lock ho sakti hai).

## GitHub par daalna
- Code GitHub par daal sakte ho, par `inventory.xlsx` aur `backups/` kabhi mat daalo. Usme customers ke naam aur phone numbers hain. `.gitignore` me ye pehle se band hain.
- Repo **private** rakho.
- GitHub khud Python server nahi chalata (GitHub Pages sirf static pages chalata hai). Online chalane ke liye PythonAnywhere jaisa hosting chahiye, aur uske pehle login/password lagana zaroori hai.

## Backup se data wapas lana (apne aap)
- Har save ke baad backup folder me `inventory-latest.xlsx` ki ek taaza copy ban jati hai. Roz ki pehli save par us din ki dated copy bhi banti hai.
- Agar `inventory.xlsx` nahi mili (naya computer, ya file delete ho gayi), to app chalu hote hi backup folder ki **sabse nayi Excel file** se data apne aap le leta hai. Kuch karna nahi padta.
- Kisi khaas Excel se data lana ho: purani `inventory.xlsx` hata do (ya naam badal do), wo Excel backup folder me rakho (usme **Items** aur **History** sheet honi chahiye), aur app dobara chalao.
- `before-reset` naam wali files is auto-restore me shamil nahi hoti, taaki reset kiya hua data wapas na aa jaye.
