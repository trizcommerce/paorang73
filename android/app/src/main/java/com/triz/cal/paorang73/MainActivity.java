package com.triz.cal.paorang73;

import android.app.Activity;
import android.app.AlertDialog;
import android.app.DownloadManager;
import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.content.IntentFilter;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.widget.Toast;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import org.json.JSONObject;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;

/** 뽕커벨님 콘텐츠 캘린더(https://trizcommerce.github.io/paorang73/)를 앱 화면으로 보여준다.
 *  내용은 사이트에서 바로 불러오므로 시트가 바뀌면 앱을 다시 설치하지 않아도 최신 내용이 보인다. */
public class MainActivity extends Activity {
    static final String HOME = "https://trizcommerce.github.io/paorang73/";
    WebView web;

    @Override
    protected void onCreate(Bundle saved) {
        super.onCreate(saved);
        web = new WebView(this);
        setContentView(web);
        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setUserAgentString(s.getUserAgentString() + " TrizCalendarApp"); // 웹 쪽에서 "앱으로 열었음"을 알아보는 표시
        s.setDomStorageEnabled(true);       // 캡션 수정본·이름 등 휴대폰 저장 기능
        s.setDatabaseEnabled(true);
        s.setLoadWithOverviewMode(true);
        s.setUseWideViewPort(true);
        s.setSupportMultipleWindows(false);
        s.setMediaPlaybackRequiresUserGesture(false);
        web.setWebChromeClient(new WebChromeClient());
        web.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView v, WebResourceRequest req) {
                Uri u = req.getUrl();
                String url = u.toString();
                // 앱 안의 화면은 그대로, 그 밖(인스타그램·구글 캘린더 구독·네이버 맞춤법 등)은 휴대폰 기본 앱으로
                if (url.startsWith(HOME) && !url.contains("/downloads/")) return false;
                try {
                    startActivity(new Intent(Intent.ACTION_VIEW, u));
                } catch (Exception ignored) { }
                return true;
            }
        });
        if (saved != null) web.restoreState(saved);
        else { web.loadUrl(HOME); checkUpdate(); }
    }

    // ---- 새 버전 알림: 아이콘·앱이 바뀌면 지우지 않고 그 위에 바로 업데이트 ----
    static final String VER_URL = HOME + "downloads/android-version.json";
    static final String APK_URL = HOME + "downloads/paorang73-calendar.apk";
    long downloadId = -1;

    long myVersion() {
        try {
            if (Build.VERSION.SDK_INT >= 28) return getPackageManager().getPackageInfo(getPackageName(), 0).getLongVersionCode();
            return getPackageManager().getPackageInfo(getPackageName(), 0).versionCode;
        } catch (Exception e) { return Long.MAX_VALUE; }
    }

    void checkUpdate() {
        new Thread(() -> {
            try {
                HttpURLConnection c = (HttpURLConnection) new URL(VER_URL + "?t=" + System.currentTimeMillis()).openConnection();
                c.setConnectTimeout(8000); c.setReadTimeout(8000); c.setUseCaches(false);
                InputStream in = c.getInputStream();
                ByteArrayOutputStream out = new ByteArrayOutputStream();
                byte[] b = new byte[4096]; int n;
                while ((n = in.read(b)) > 0) out.write(b, 0, n);
                in.close();
                long latest = new JSONObject(out.toString("UTF-8")).optLong("versionCode", 0);
                if (latest > myVersion()) runOnUiThread(this::askUpdate);
            } catch (Exception ignored) { }
        }).start();
    }

    void askUpdate() {
        if (isFinishing()) return;
        new AlertDialog.Builder(this)
            .setTitle("새 버전이 있어요")
            .setMessage("앱 아이콘과 기능이 새로워졌어요. 지금 업데이트할까요?\n앱을 지우지 않고 그대로 덮어써서, 저장해 둔 내용은 그대로 남아요.")
            .setPositiveButton("업데이트", (d, w) -> startDownload())
            .setNegativeButton("나중에", null)
            .show();
    }

    void startDownload() {
        try {
            DownloadManager dm = (DownloadManager) getSystemService(Context.DOWNLOAD_SERVICE);
            DownloadManager.Request r = new DownloadManager.Request(Uri.parse(APK_URL))
                .setTitle("뽕커벨 캘린더 업데이트")
                .setMimeType("application/vnd.android.package-archive")
                .setNotificationVisibility(DownloadManager.Request.VISIBILITY_VISIBLE_NOTIFY_COMPLETED)
                .setDestinationInExternalFilesDir(this, Environment.DIRECTORY_DOWNLOADS, "paorang73-calendar-" + System.currentTimeMillis() + ".apk");
            downloadId = dm.enqueue(r);
            BroadcastReceiver done = new BroadcastReceiver() {
                @Override public void onReceive(Context ctx, Intent i) {
                    if (i.getLongExtra(DownloadManager.EXTRA_DOWNLOAD_ID, -2) != downloadId) return;
                    try { unregisterReceiver(this); } catch (Exception ignored) { }
                    Uri apk = dm.getUriForDownloadedFile(downloadId);
                    if (apk == null) { openInBrowser(); return; }
                    Intent install = new Intent(Intent.ACTION_VIEW)
                        .setDataAndType(apk, "application/vnd.android.package-archive")
                        .addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_ACTIVITY_NEW_TASK);
                    try { startActivity(install); } catch (Exception e) { openInBrowser(); }
                }
            };
            IntentFilter f = new IntentFilter(DownloadManager.ACTION_DOWNLOAD_COMPLETE);
            if (Build.VERSION.SDK_INT >= 33) registerReceiver(done, f, Context.RECEIVER_EXPORTED);
            else registerReceiver(done, f);
            Toast.makeText(this, "새 버전을 내려받고 있어요. 다 받으면 설치 화면이 열려요.", Toast.LENGTH_LONG).show();
        } catch (Exception e) { openInBrowser(); }
    }

    void openInBrowser() {
        try { startActivity(new Intent(Intent.ACTION_VIEW, Uri.parse(APK_URL))); } catch (Exception ignored) { }
    }

    @Override
    protected void onSaveInstanceState(Bundle out) {
        super.onSaveInstanceState(out);
        web.saveState(out);
    }

    @Override
    public void onBackPressed() {
        if (web.canGoBack()) web.goBack();
        else super.onBackPressed();
    }
}
