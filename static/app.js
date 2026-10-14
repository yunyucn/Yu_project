/* 两个页面共用一份脚本：按页面上有没有对应元素来决定跑哪一段。 */
(function () {
  "use strict";

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c];
    });
  }

  function readResponse(response) {
    return response.text().then(function (text) {
      var data;
      try { data = JSON.parse(text); }
      catch (error) {
        throw new Error(response.status === 413 ? '文件总大小不能超过 300 MB。' : '服务暂时不可用，请稍后重试。');
      }
      if (!response.ok) {
        var reference = data.task_id || data.request_id;
        throw new Error((data.error || '操作未完成，请重试。') + (reference ? '（编号：' + reference + '）' : ''));
      }
      return data;
    });
  }

  /* ---------------- 首页：选择文件 + 上传 ---------------- */
  var uploadForm = document.getElementById("uploadForm");
  if (uploadForm) {
    var input = document.getElementById("fileInput");
    var drop = document.getElementById("drop");
    var list = document.getElementById("fileList");
    var msg = document.getElementById("uploadMsg");
    var btn = document.getElementById("uploadBtn");
    var picked = [];

    function render() {
      list.innerHTML = picked.map(function (f, index) {
        return "<li>" + esc(f.name) + " <span class='hint'>(" +
               (f.size / 1048576).toFixed(1) + " MB)</span> <button type='button' class='remove-file' data-index='" + index + "'>移除</button></li>";
      }).join("");
    }
    function addFiles(fl) {
      if (btn.disabled) return;
      for (var i = 0; i < fl.length; i++) {
        var file = fl[i];
        if (!file.name.toLowerCase().endsWith('.pdf')) continue;
        if (!picked.some(function (existing) {
          return existing.name === file.name && existing.size === file.size && existing.lastModified === file.lastModified;
        })) picked.push(file);
      }
      render();
    }
    list.addEventListener('click', function (event) {
      if (btn.disabled || !event.target.matches('button[data-index]')) return;
      picked.splice(Number(event.target.dataset.index), 1);
      render();
    });
    input.addEventListener("change", function () { addFiles(input.files); input.value = ""; });
    ["dragenter", "dragover"].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add("over"); });
    });
    ["dragleave", "drop"].forEach(function (ev) {
      drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.remove("over"); });
    });
    drop.addEventListener("drop", function (e) {
      if (e.dataTransfer && e.dataTransfer.files) addFiles(e.dataTransfer.files);
    });

    uploadForm.addEventListener("submit", function (e) {
      e.preventDefault();
      if (btn.disabled) return;
      if (!picked.length) {
        msg.innerHTML = "<span class='alert alert-warn'>先选至少一份 PDF。</span>";
        return;
      }
      if (picked.reduce(function (size, file) { return size + file.size; }, 0) > 300 * 1048576) {
        msg.textContent = '文件总大小不能超过 300 MB。';
        return;
      }
      var fd = new FormData();
      picked.forEach(function (f) { fd.append("files", f); });
      btn.disabled = true;
      msg.innerHTML = "<span class='hint'>正在读取财报…</span>";
      fetch("/api/tasks", { method: "POST", body: fd })
        .then(readResponse)
        .then(function (res) {
          window.location.href = res.redirect;
        })
        .catch(function (err) {
          btn.disabled = false;
          msg.innerHTML = "<span class='alert alert-error'>" + esc(err.message) + "</span>";
        });
    });
  }

  /* ---------------- 任务页 ---------------- */
  var body = document.body;
  var startBtn = document.getElementById("startBtn");
  if (startBtn && body.dataset.task) {
    var taskId = body.dataset.task;
    var bar = document.getElementById("bar");
    var stepText = document.getElementById("stepText");
    var progressWrap = document.getElementById("progressWrap");
    var startMsg = document.getElementById("startMsg");
    var resultCard = document.getElementById("resultCard");
    var resultMessage = document.getElementById("resultMessage");
    var downloadLink = document.getElementById("downloadLink");
    var timer = null;
    var finished = false;
    var generation = 0;
    var polling = false;
    var blocked = startBtn.disabled;
    var sawRunning = false;   /* 这次打开页面时任务在跑，跑完才自动滚到结果 */

    function paint(t) {
      if (t.status === "uploaded") { progressWrap.style.display = "none"; return; }
      var total = t.total_steps || 11;
      var step = t.step || 0;
      var pct = Math.max(3, Math.round(100 * step / total));
      if (t.status === "done" || t.status === 'review') pct = 100;
      progressWrap.style.display = "block";
      bar.style.width = pct + "%";
      stepText.textContent = "第 " + step + "/" + total + " 步 · " + (t.step_name || t.status);
    }

    function finish(t) {
      if (finished) return;
      finished = true;
      if (timer) { clearInterval(timer); timer = null; }
      paint(t);
      resultCard.style.display = "block";
      startBtn.disabled = blocked;
      startBtn.textContent = '重新生成';
      startMsg.innerHTML = "";
      resultMessage.textContent = t.message || (t.has_model ? '已有文件可下载。' : '没有生成文件。');
      if (t.status === 'failed' || t.status === 'stopped') {
        resultMessage.textContent += ' 任务编号：' + taskId;
      }
      if (t.has_model) {
        var downloadBtn = document.getElementById('downloadBtn');
        if (downloadBtn) downloadBtn.textContent = t.download_label || '下载模型文件';
        downloadLink.style.display = "inline-block";
        downloadLink.href = "/api/tasks/" + taskId + "/download";
      } else {
        downloadLink.style.display = "none";
      }
      // 结果在页面最底下。只在「这次眼看着它跑完」时滚过去：重新打开一个早就
      // 跑完的任务，用户多半是来看上面那张表的，自动滚到底会把人甩走。
      if (sawRunning && resultCard.scrollIntoView) {
        resultCard.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    }

    function poll() {
      if (polling) return;
      polling = true;
      var currentGeneration = generation;
      var controller = new AbortController();
      var requestTimer = setTimeout(function () { controller.abort(); }, 15000);
      fetch("/api/tasks/" + taskId, { signal: controller.signal }).then(readResponse).then(function (t) {
        if (currentGeneration !== generation) return;
        startMsg.textContent = '';
        paint(t);
        if (t.status === "done" || t.status === 'review' || t.status === "stopped" || t.status === "failed") {
          finish(t);
        } else if (t.status === "queued" || t.status === "running" || t.status === 'uploading') {
          finished = false;
          resultCard.style.display = 'none';
          startBtn.disabled = true;
          sawRunning = true;
          if (!timer) timer = setInterval(poll, 2000);
        } else {
          startBtn.disabled = blocked;
          if (timer) { clearInterval(timer); timer = null; }
        }
      }).catch(function () {
        if (currentGeneration !== generation) return;
        startMsg.textContent = '连接暂时中断，正在重试…';
        if (!timer) timer = setInterval(poll, 5000);
      }).finally(function () {
        clearTimeout(requestTimer);
        polling = false;
      });
    }

    startBtn.addEventListener("click", function () {
      if (startBtn.disabled) return;
      if (!window.confirm("开始生成估值模型？")) return;
      var val = function (id) {
        var el = document.getElementById(id);
        return el ? el.value : "";
      };
      var payload = {
        company_name: val("company_name"),
        company_code: val("company_code"),
        report_period: val("report_period"),
        total_shares_mn: val("total_shares_mn"),
        float_shares_mn: val("float_shares_mn"),
        allow_incomplete_history: true
      };
      startBtn.disabled = true;
      generation += 1;
      finished = false;
      sawRunning = true;
      resultCard.style.display = 'none';
      downloadLink.style.display = 'none';
      downloadLink.removeAttribute('href');
      if (timer) { clearInterval(timer); timer = null; }
      startMsg.innerHTML = "<span class='hint'>正在提交…</span>";
      fetch("/api/tasks/" + taskId + "/start", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      }).then(readResponse)
        .then(function (res) {
          progressWrap.style.display = "block";
          if (!timer) timer = setInterval(poll, 2000);
          poll();
        })
        .catch(function (err) {
          startBtn.disabled = blocked;
          startMsg.innerHTML = "<span class='alert alert-error'>" + esc(err.message) + "</span>";
          if (!timer) timer = setInterval(poll, 5000);
        });
    });

    poll();
  }
})();
