# ExcelBot Pro - Android App User Manual 📱

**Complete Guide to Installing and Using ExcelBot Pro on Android**

---

## 📱 What is This?

ExcelBot Pro for Android is a mobile version of the Excel VBA macro generator. It connects to a cloud-hosted version of the application, allowing you to generate VBA macros directly from your Android phone or tablet!

---

## 🎯 Two Ways to Use ExcelBot Pro on Android

### Option 1: Pre-Built APK (Easiest)
Download and install the ready-made Android app (see installation steps below)

### Option 2: Build from Source
Build the app yourself using Android Studio (for developers)

---

## 📦 Option 1: Installing the Pre-Built APK

### What You Need:
- Android phone or tablet (Android 5.0 or higher)
- 10 MB of free space
- Internet connection

### Step-by-Step Installation:

#### Step 1: Enable Installation from Unknown Sources

**For Android 8.0 and above:**
1. Open **Settings** on your Android device
2. Tap **Apps & notifications** (or **Apps**)
3. Tap **Advanced** → **Special app access**
4. Tap **Install unknown apps**
5. Select your browser (e.g., Chrome)
6. Enable **Allow from this source**

**For Android 7.1 and below:**
1. Open **Settings**
2. Tap **Security**
3. Enable **Unknown sources**
4. Tap **OK** to confirm

#### Step 2: Download the APK

1. Download `ExcelBot-Pro-Android.apk` to your device
2. You can find it in your **Downloads** folder

#### Step 3: Install the App

1. Open your file manager
2. Navigate to **Downloads**
3. Tap on `ExcelBot-Pro-Android.apk`
4. Tap **Install**
5. Wait for installation to complete
6. Tap **Open** to launch the app

#### Step 4: First Launch

When you first open the app, it will try to connect to the default server. You need to configure your own server URL (see Cloud Deployment section below).

---

## ☁️ Cloud Deployment (Required for Android App)

The Android app needs to connect to a cloud-hosted version of ExcelBot Pro. Here's how to set it up for FREE:

### Deploying to Hugging Face Spaces (FREE)

**Why Hugging Face?**
- Completely FREE hosting
- Perfect for Gradio apps
- Always online
- No credit card required

#### Step 1: Create a Hugging Face Account

1. Go to https://huggingface.co
2. Click **Sign Up** in the top-right corner
3. Fill in:
   - Email address
   - Password
   - Username (remember this!)
4. Verify your email
5. Complete profile setup

#### Step 2: Create a New Space

1. Log in to Hugging Face
2. Go to https://huggingface.co/spaces
3. Click **Create new Space** button
4. Fill in the form:
   - **Space name**: `excelbot-pro` (or any name you prefer)
   - **License**: Select "MIT"
   - **Select the Space SDK**: Choose **Gradio**
   - **Space hardware**: Select "CPU basic" (FREE tier)
   - **Visibility**: Choose "Public" (free) or "Private" (paid)
5. Click **Create Space**

#### Step 3: Upload Application Files

You'll see an empty Space. Now upload the files:

1. In your Space, click the **Files** tab
2. Click **Add file** → **Upload files**
3. Upload these files from the `cloud-deployment` folder:
   - `app.py`
   - `requirements.txt`
   - `README.md`
4. Click **Commit changes to main**

#### Step 4: Wait for Deployment

1. The Space will automatically build and deploy
2. You'll see a "Building..." status
3. Wait 2-5 minutes for the build to complete
4. Once ready, you'll see "Running" status

#### Step 5: Get Your App URL

Your app is now live at:
```
https://huggingface.co/spaces/YOUR_USERNAME/excelbot-pro
```

Example: If your username is "john123", your URL would be:
```
https://huggingface.co/spaces/john123/excelbot-pro
```

**Save this URL - you'll need it for the Android app!**

---

## ⚙️ Configuring the Android App

### Setting Your Server URL

1. Open the ExcelBot Pro app on your Android device
2. Tap the **three dots** (⋮) in the top-right corner
3. Tap **Settings**
4. In the **Server URL** field, enter your Hugging Face Space URL:
   ```
   https://huggingface.co/spaces/YOUR_USERNAME/excelbot-pro
   ```
5. Tap **Save Settings**
6. Restart the app
7. The app will now connect to your cloud-hosted ExcelBot Pro!

---

## 🎮 Using the Android App

### Generating VBA Macros

1. Open the app
2. Wait for it to load (you'll see the ExcelBot Pro interface)
3. In the text field, type what you want:
   - "format my data"
   - "sort by first column"
   - "create a chart"
   - "calculate totals"
4. Tap **Generate VBA Macro**
5. The VBA code will appear below
6. **Long press** on the code to copy it
7. Paste it into Excel on your computer

### Analyzing Excel Files

1. Tap the **upload** area
2. Select an Excel file from your device
3. The app will show:
   - Number of rows
   - Number of columns
   - Column names

### Pushing to GitHub (Optional)

If you configured GitHub:
1. Generate a VBA macro first
2. Enter your repository name: `username/repo-name`
3. Enter a file name: `my_macro.vba`
4. Tap **Push to GitHub**

---

## 🔧 Option 2: Building from Source (Developers)

### What You Need:
- Android Studio (latest version)
- Java Development Kit (JDK) 8 or higher
- Android SDK

### Step-by-Step Build Process:

#### Step 1: Install Android Studio

1. Download from: https://developer.android.com/studio
2. Run the installer
3. Follow the setup wizard
4. Install Android SDK (API 33 recommended)

#### Step 2: Open the Project

1. Extract the `android-app` folder
2. Open Android Studio
3. Click **File** → **Open**
4. Navigate to the `android-app` folder
5. Click **OK**
6. Wait for Gradle sync to complete

#### Step 3: Configure the Project

1. Open `MainActivity.java`
2. Find this line:
   ```java
   private static final String DEFAULT_SERVER_URL = "https://huggingface.co/spaces/YOUR_USERNAME/excelbot-pro";
   ```
3. Replace `YOUR_USERNAME` with your actual Hugging Face username
4. Save the file

#### Step 4: Build the APK

**For Testing (Debug Build):**
1. Click **Build** → **Build Bundle(s) / APK(s)** → **Build APK(s)**
2. Wait for the build to complete
3. Click **locate** in the notification
4. Find `app-debug.apk`

**For Release (Signed APK):**
1. Click **Build** → **Generate Signed Bundle / APK**
2. Select **APK**
3. Click **Next**
4. Create a new keystore or use existing
5. Fill in keystore details
6. Click **Next**
7. Select **release** build variant
8. Click **Finish**
9. Find your APK in `app/release/`

#### Step 5: Install on Your Device

1. Enable USB debugging on your Android device:
   - Go to **Settings** → **About phone**
   - Tap **Build number** 7 times
   - Go back to **Settings** → **Developer options**
   - Enable **USB debugging**
2. Connect your device to your computer
3. In Android Studio, click the **Run** button (▶)
4. Select your device
5. The app will install and launch

---

## 🔑 Configuring GitHub Token (Optional)

To use GitHub features with your cloud deployment:

### Step 1: Generate GitHub Token

Follow the same steps as in the main USER_MANUAL.md to create a Personal Access Token.

### Step 2: Add to Hugging Face Space

1. Go to your Space on Hugging Face
2. Click **Settings** tab
3. Scroll down to **Repository secrets**
4. Click **New secret**
5. Name: `GITHUB_TOKEN`
6. Value: Paste your GitHub token
7. Click **Add**
8. Restart your Space (click Settings → Factory reboot)

Now the GitHub push feature will work in your Android app!

---

## 📊 App Features

### Main Screen
- **WebView**: Displays the full ExcelBot Pro interface
- **Progress Bar**: Shows loading status
- **Menu Button** (⋮): Access settings and about info

### Menu Options
- **Refresh**: Reload the application
- **Settings**: Configure server URL
- **About**: App information

### Settings Screen
- **Server URL Input**: Enter your Hugging Face Space URL
- **Save Button**: Save configuration
- **Reset Button**: Restore default URL

---

## ❓ Troubleshooting

### Issue 1: App Won't Install

**Error: "App not installed"**

**Solutions:**
1. Make sure you enabled "Unknown sources" (see installation steps)
2. Delete any old version of the app first
3. Clear cache: Settings → Apps → Package Installer → Clear cache
4. Try downloading the APK again

---

### Issue 2: Can't Connect to Server

**Error: "Error loading page"**

**Solutions:**
1. Check your internet connection
2. Verify your server URL in Settings:
   - It should start with `https://`
   - No trailing slash at the end
   - Username is correct
3. Make sure your Hugging Face Space is running:
   - Visit the URL in your phone's browser first
   - If it loads there, it should work in the app
4. Try **Refresh** from the menu

---

### Issue 3: App is Slow

**Possible causes:**
1. Slow internet connection
2. Hugging Face Space is cold-starting (first load after inactivity)

**Solutions:**
1. Wait 30-60 seconds for initial load
2. Use a faster WiFi connection
3. The app will be faster on subsequent uses

---

### Issue 4: Can't Upload Excel Files

**Problem**: File upload doesn't work

**Solution:**
This is a known limitation of the WebView approach. To upload files:
1. Use the desktop/PC version instead, OR
2. Upload files directly to your Hugging Face Space and reference them

---

### Issue 5: GitHub Push Doesn't Work

**Error: "GitHub token not configured"**

**Solution:**
1. Add your GitHub token to Hugging Face Space secrets
2. See "Configuring GitHub Token" section above
3. Restart your Space after adding the token

---

## 🔒 Security & Privacy

### Data Privacy
- All data stays between your device and your Hugging Face Space
- If using Public Space: Others can access your deployment URL
- If using Private Space: Only you can access (requires paid plan)

### Best Practices
1. **Don't upload sensitive files** to public Spaces
2. **Use Private Spaces** for confidential work
3. **Keep GitHub tokens secret** - never share them
4. **Review VBA code** before running in Excel

---

## 📱 System Requirements

### Minimum Requirements:
- Android 5.0 (Lollipop) or higher
- 50 MB free storage
- Internet connection (WiFi or mobile data)

### Recommended:
- Android 8.0 or higher
- 100 MB free storage
- WiFi connection for faster loading

---

## 🆚 Android vs Desktop Version

| Feature | Android | Desktop |
|---------|---------|---------|
| VBA Generation | ✅ Yes | ✅ Yes |
| Excel Analysis | ⚠️ Limited | ✅ Yes |
| GitHub Push | ✅ Yes | ✅ Yes |
| File Upload | ⚠️ Limited | ✅ Yes |
| Offline Use | ❌ No | ✅ Yes |
| Portability | ✅ High | ⚠️ Medium |

**Recommendation:** Use Android for quick VBA generation on-the-go, and desktop for full features.

---

## 🌟 Pro Tips

### Tip 1: Bookmark Your Space
Save your Hugging Face Space URL in your phone's browser for quick access without the app.

### Tip 2: Create Shortcuts
Add common VBA tasks as notes in your phone for quick copy-paste into the app.

### Tip 3: Share with Team
Share your Space URL with colleagues so they can use ExcelBot Pro too!

### Tip 4: Keep App Updated
When new features are added to the cloud version, they automatically appear in your app!

---

## 📞 Additional Resources

### Learning & Support
- **Hugging Face Docs**: https://huggingface.co/docs/hub/spaces
- **Android Studio**: https://developer.android.com/studio/intro
- **Gradio Documentation**: https://www.gradio.app/docs/

### Video Tutorials
- Installing APK on Android: Search YouTube for "how to install APK on Android"
- Creating Hugging Face Space: Search for "Hugging Face Spaces tutorial"

---

## ✅ Quick Start Checklist

- [ ] Android device with version 5.0+
- [ ] Enabled installation from unknown sources
- [ ] Created Hugging Face account
- [ ] Created Space on Hugging Face
- [ ] Uploaded app files to Space
- [ ] Space is running and accessible
- [ ] Noted down Space URL
- [ ] Installed ExcelBot Pro APK on phone
- [ ] Configured server URL in app settings
- [ ] Successfully generated first VBA macro!

---

## 🎉 Summary

You now have ExcelBot Pro running on your Android device! The app connects to your free cloud deployment, giving you access to powerful VBA generation anywhere, anytime.

**Key Points:**
- ✅ 100% FREE to use (Hugging Face + Android app)
- ✅ No coding knowledge required
- ✅ Generate VBA macros on-the-go
- ✅ Cloud-based, always up-to-date
- ✅ Works on any Android 5.0+ device

**Enjoy automating Excel from your phone!** 📱✨

---

*Version 1.0 - Android Edition*
*For technical support, refer to the troubleshooting section or visit the Hugging Face documentation.*
