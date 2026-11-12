package az.asf.workerassistant;

import android.app.Activity;
import android.content.ContentValues;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.webkit.URLUtil;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.Toast;

import com.chaquo.python.Python;
import com.chaquo.python.android.AndroidPlatform;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;

/**
 * Server telefonun özündədir: Python (Chaquopy) wa_android.basla() ilə GitHub-dan son versiyanı çəkib
 * 127.0.0.1:8000-də qaldırır, interfeys WebView-da açılır. Fayl seçmək (📎) və hazır Excel-ləri
 * «Yükləmələr»ə saxlamaq dəstəklənir.
 */
public class MainActivity extends Activity {
    private static final int PORT = 8000;
    private static final String UNVAN = "http://127.0.0.1:" + PORT + "/";
    private static final int FAYL_SEC = 1;
    private WebView web;
    private ValueCallback<Uri[]> faylCavabi;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        web = new WebView(this);
        WebSettings ws = web.getSettings();
        ws.setJavaScriptEnabled(true);
        ws.setDomStorageEnabled(true);
        ws.setAllowFileAccess(false);
        web.setWebViewClient(new WebViewClient());
        web.setWebChromeClient(new WebChromeClient() {
            @Override
            public boolean onShowFileChooser(WebView v, ValueCallback<Uri[]> cb, FileChooserParams p) {
                if (faylCavabi != null) faylCavabi.onReceiveValue(null);
                faylCavabi = cb;
                Intent i = p.createIntent();
                i.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, p.getMode() == FileChooserParams.MODE_OPEN_MULTIPLE);
                try {
                    startActivityForResult(i, FAYL_SEC);
                } catch (Exception e) {
                    faylCavabi = null;
                    return false;
                }
                return true;
            }
        });
        web.setDownloadListener((url, userAgent, disposition, mime, length) -> yukle(url, disposition, mime));
        setContentView(web);
        mesaj("Worker's Assistant başladılır…", "GitHub-dan son versiya yoxlanılır və server telefonda qaldırılır. İlk dəfə bir az çəkə bilər.");

        new Thread(() -> {
            try {
                if (!Python.isStarted()) Python.start(new AndroidPlatform(this));
                String info = Python.getInstance().getModule("wa_android")
                        .callAttr("basla", getFilesDir().getAbsolutePath(), PORT).toString();
                hazirOlanaQederGozle();
                runOnUiThread(() -> {
                    Toast.makeText(this, info.contains("GitHub (") ? "Son versiya: GitHub" : "Versiya: " + menbe(info), Toast.LENGTH_LONG).show();
                    web.loadUrl(UNVAN);
                });
            } catch (Throwable t) {
                runOnUiThread(() -> mesaj("Server qalxmadı", String.valueOf(t)));
            }
        }).start();
    }

    private static String menbe(String json) {
        int i = json.indexOf("\"menbe\": \"");
        if (i < 0) return "";
        int b = i + 10, s = json.indexOf('"', b);
        return s > b ? json.substring(b, s) : "";
    }

    private void hazirOlanaQederGozle() throws InterruptedException {
        for (int i = 0; i < 300; i++) {
            try {
                HttpURLConnection c = (HttpURLConnection) new URL(UNVAN + "saglamliq").openConnection();
                c.setConnectTimeout(500);
                if (c.getResponseCode() == 200) return;
            } catch (Exception ignored) {
            }
            Thread.sleep(200);
        }
    }

    private void mesaj(String basliq, String metn) {
        String html = "<html><body style='background:#050505;color:#e8e8ea;font-family:sans-serif;padding:32px'>"
                + "<h3 style='color:#ff2a2a;letter-spacing:.1em'>" + basliq + "</h3><p style='color:#9a8f8f'>"
                + metn.replace("<", "&lt;") + "</p></body></html>";
        web.loadDataWithBaseURL(null, html, "text/html", "utf-8", null);
    }

    /** Hazır Excel-i telefonun «Yükləmələr» qovluğuna saxlayır (server telefonun özündə olduğu üçün birbaşa oxunur). */
    private void yukle(String url, String disposition, String mime) {
        String ad = URLUtil.guessFileName(url, disposition, mime);
        new Thread(() -> {
            try {
                HttpURLConnection c = (HttpURLConnection) new URL(url).openConnection();
                OutputStream out;
                Uri uri = null;
                if (Build.VERSION.SDK_INT >= 29) {
                    ContentValues cv = new ContentValues();
                    cv.put(MediaStore.Downloads.DISPLAY_NAME, ad);
                    cv.put(MediaStore.Downloads.MIME_TYPE, mime != null ? mime : "application/octet-stream");
                    uri = getContentResolver().insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, cv);
                    out = getContentResolver().openOutputStream(uri);
                } else {
                    File q = getExternalFilesDir(Environment.DIRECTORY_DOWNLOADS);
                    out = new FileOutputStream(new File(q, ad));
                }
                try (InputStream in = c.getInputStream(); OutputStream o = out) {
                    byte[] buf = new byte[8192];
                    int n;
                    while ((n = in.read(buf)) > 0) o.write(buf, 0, n);
                }
                final Uri ac = uri;
                runOnUiThread(() -> {
                    Toast.makeText(this, "Yükləndi: " + ad + " (Yükləmələr)", Toast.LENGTH_LONG).show();
                    if (ac != null) {
                        try {
                            Intent i = new Intent(Intent.ACTION_VIEW).setDataAndType(ac, mime).addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                            startActivity(i);
                        } catch (Exception ignored) {  // Excel açan tətbiq yoxdursa — fayl Yükləmələrdə qalır
                        }
                    }
                });
            } catch (Exception e) {
                runOnUiThread(() -> Toast.makeText(this, "Yüklənmədi: " + e.getMessage(), Toast.LENGTH_LONG).show());
            }
        }).start();
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        if (requestCode == FAYL_SEC && faylCavabi != null) {
            Uri[] secilen = null;
            if (resultCode == RESULT_OK && data != null) {
                if (data.getClipData() != null) {  // bir neçə fayl (çıxarış + şablon + müqavilələr)
                    secilen = new Uri[data.getClipData().getItemCount()];
                    for (int i = 0; i < secilen.length; i++) secilen[i] = data.getClipData().getItemAt(i).getUri();
                } else if (data.getData() != null) {
                    secilen = new Uri[]{data.getData()};
                }
            }
            faylCavabi.onReceiveValue(secilen);
            faylCavabi = null;
            return;
        }
        super.onActivityResult(requestCode, resultCode, data);
    }

    @Override
    public void onBackPressed() {
        if (web.canGoBack()) web.goBack();
        else super.onBackPressed();
    }
}
