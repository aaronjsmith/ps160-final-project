(function(){
  function escapeHtml(str){
    return String(str)
      .replace(/&/g,'&amp;')
      .replace(/</g,'&lt;')
      .replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;')
      .replace(/'/g,'&#039;');
  }

  function renderSimpleArticle(container, content){
    var html = '';
    if (content.title) html += '<h1>' + escapeHtml(content.title) + '</h1>';
    if (content.intro) html += '<p>' + escapeHtml(content.intro) + '</p>';
    if (Array.isArray(content.sections)){
      content.sections.forEach(function(section){
        if (section.heading) html += '<h2>' + escapeHtml(section.heading) + '</h2>';
        if (section.quote) html += '<blockquote>' + escapeHtml(section.quote) + '</blockquote>';
        if (section.body) html += '<p>' + escapeHtml(section.body) + '</p>';
        if (Array.isArray(section.list)){
          html += '<ul>' + section.list.map(function(item){ return '<li>' + escapeHtml(item) + '</li>'; }).join('') + '</ul>';
        }
      });
    }
    container.innerHTML = html;
  }

  function renderHome(){
    var home = window.__CONTENT__.home;
    if (!home) return;

    var titleEl = document.getElementById('home-title');
    var introEl = document.getElementById('home-intro');
    var bulletsEl = document.getElementById('home-bullets');
    var captionEl = document.getElementById('home-image-caption');
    var sectionsEl = document.getElementById('home-sections');

    if (titleEl) titleEl.textContent = home.title || titleEl.textContent;
    if (introEl) introEl.textContent = home.intro || introEl.textContent;
    if (Array.isArray(home.bullets) && bulletsEl){
      bulletsEl.innerHTML = home.bullets.map(function(b){
        return '<li class="flex items-start gap-2"><span class="mt-1 h-1.5 w-1.5 rounded-full bg-amber-700"></span><span>' + escapeHtml(b) + '</span></li>';
      }).join('');
    }
    if (captionEl && home.imageCaption) captionEl.textContent = home.imageCaption;
    if (Array.isArray(home.sections) && sectionsEl){
      sectionsEl.querySelectorAll('[data-section-index]').forEach(function(secEl){
        var idx = parseInt(secEl.getAttribute('data-section-index'), 10);
        var sec = home.sections[idx];
        if (!sec) return;
        var h = secEl.querySelector('[data-section-heading]');
        var p = secEl.querySelector('[data-section-body]');
        if (h && sec.heading) h.textContent = sec.heading;
        if (p && sec.body) p.textContent = sec.body;
      });
    }
  }

  function loadContent(){
    fetch('assets/content.json')
      .then(function(r){ return r.json(); })
      .then(function(json){
        window.__CONTENT__ = json || {};

        var article = document.querySelector('[data-content-key]');
        if (article){
          var key = article.getAttribute('data-content-key');
          var content = window.__CONTENT__[key];
          if (content){ renderSimpleArticle(article, content); }
        }

        // optional home wiring
        renderHome();
      })
      .catch(function(){ /* silent */ });
  }

  if (document.readyState === 'complete' || document.readyState === 'interactive'){
    loadContent();
  } else {
    document.addEventListener('DOMContentLoaded', loadContent);
  }
})();


