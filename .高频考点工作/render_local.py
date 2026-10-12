# Bundled LO and Poppler require GLIBC 2.38; installed binaries support this host.
import importlib.util,os
spec=importlib.util.spec_from_file_location('docx_renderer','/home/zyf/.codex/plugins/cache/openai-primary-runtime/documents/26.1011.11113/skills/documents/render_docx.py')
mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
def installed_soffice():
 os.environ['PATH']='/usr/bin:/bin:'+os.environ.get('PATH','')
 return '/usr/bin/soffice'
mod._resolve_soffice=installed_soffice
mod.main()
