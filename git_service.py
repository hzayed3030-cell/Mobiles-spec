# git_service.py
import subprocess
import os
import shutil
from typing import Dict, Any, Tuple, Optional, List

DEFAULT_PROJECT_GITHUB_URL = "https://github.com/hzayed3030-cell/Mobiles-spec.git"
DEFAULT_PROJECT_WEB_URL = "https://github.com/hzayed3030-cell/Mobiles-spec"

def run_git_cmd(args: list, cwd: Optional[str] = None) -> Tuple[bool, str]:
    """تنفيذ أمر git وإرجاع (success, output/error)"""
    if cwd is None:
        cwd = os.path.dirname(os.path.abspath(__file__))
    try:
        res = subprocess.run(
            ["git"] + args,
            cwd=cwd,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            shell=False
        )
        out = (res.stdout or "") + (res.stderr or "")
        return (res.returncode == 0, out.strip())
    except Exception as e:
        return (False, f"خطأ أثناء تشغيل Git: {str(e)}")

def check_large_build_dirs(cwd: Optional[str] = None) -> List[Dict[str, Any]]:
    """فحص وجود فهارس build أو dist الخاصة بحزم ملفات EXE ذات الحجم الكبير"""
    if cwd is None:
        cwd = os.path.dirname(os.path.abspath(__file__))
    found = []
    try:
        for item in os.listdir(cwd):
            full_path = os.path.join(cwd, item)
            if os.path.isdir(full_path) and item.lower() in ["build", "dist"]:
                files_count = 0
                total_size = 0
                for dirpath, dirnames, filenames in os.walk(full_path):
                    for filename in filenames:
                        files_count += 1
                        total_size += os.path.getsize(os.path.join(dirpath, filename))
                
                # إذا كان المجلد فارغاً تماماً نقوم بإزالته تلقائياً
                if files_count == 0:
                    try:
                        shutil.rmtree(full_path, ignore_errors=True)
                    except Exception:
                        pass
                    continue
                    
                size_mb = round(total_size / (1024 * 1024), 2)
                if size_mb >= 0.5 or files_count > 5:
                    found.append({
                        "name": item,
                        "path": full_path,
                        "size_mb": size_mb,
                        "files_count": files_count
                    })
    except Exception:
        pass
    return found

def delete_build_dirs(cwd: Optional[str] = None) -> Tuple[bool, str]:
    """حذف مجلدات build و dist الخاصة بـ EXE وتحديث .gitignore لمنع رفعها"""
    if cwd is None:
        cwd = os.path.dirname(os.path.abspath(__file__))
    deleted = []
    try:
        for item in os.listdir(cwd):
            full_path = os.path.join(cwd, item)
            if os.path.isdir(full_path) and item.lower() in ["build", "dist"]:
                shutil.rmtree(full_path, ignore_errors=True)
                deleted.append(item)
        
        # التأكد من استبعادها في .gitignore
        gitignore_path = os.path.join(cwd, ".gitignore")
        current_content = ""
        if os.path.exists(gitignore_path):
            with open(gitignore_path, "r", encoding="utf-8") as f:
                current_content = f.read()
        
        needed = ["build/", "dist/", "*.spec", "*.exe"]
        added = [n for n in needed if n not in current_content]
        if added:
            with open(gitignore_path, "a", encoding="utf-8") as f:
                f.write("\n# PyInstaller & EXE Packaging\n" + "\n".join(added) + "\n")
                
        return (True, f"تم حذف الفهارس بنجاح: {', '.join(deleted)}" if deleted else "لم يتم العثور على فهارس build أو dist")
    except Exception as e:
        return (False, f"حدث خطأ أثناء الحذف: {str(e)}")

def get_git_info(cwd: Optional[str] = None) -> Dict[str, Any]:
    """استرجاع معلومات حالة Git الحالية"""
    if cwd is None:
        cwd = os.path.dirname(os.path.abspath(__file__))
    
    is_repo, _ = run_git_cmd(["rev-parse", "--is-inside-work-tree"], cwd=cwd)
    if not is_repo:
        return {
            "is_repo": False,
            "branch": "main",
            "remote": DEFAULT_PROJECT_GITHUB_URL,
            "status_text": "المستودع المحلي غير مهيأ بعد",
            "has_changes": False,
            "last_commit": "-"
        }
    
    _, branch = run_git_cmd(["branch", "--show-current"], cwd=cwd)
    s_rem, remote = run_git_cmd(["remote", "get-url", "origin"], cwd=cwd)
    _, status = run_git_cmd(["status", "--short"], cwd=cwd)
    _, last_c = run_git_cmd(["log", "-1", "--pretty=format:%h - %s (%cr)"], cwd=cwd)
    
    clean_remote = remote if (s_rem and remote and not remote.startswith("error")) else DEFAULT_PROJECT_GITHUB_URL
    
    return {
        "is_repo": True,
        "branch": branch if branch else "main",
        "remote": clean_remote,
        "status_text": status if status else "المستودع نظيف ومحدث ومطابق للسحابة",
        "has_changes": bool(status.strip()),
        "last_commit": last_c if last_c else "-"
    }

def get_git_status_summary(cwd: Optional[str] = None) -> Dict[str, Any]:
    """استرجاع ملخص مفصل للملفات المعدلة وحالة المزامنة"""
    if cwd is None:
        cwd = os.path.dirname(os.path.abspath(__file__))
    
    is_repo, _ = run_git_cmd(["rev-parse", "--is-inside-work-tree"], cwd=cwd)
    if not is_repo:
        return {"is_repo": False, "files": [], "total_changes": 0, "status_clean": True}
        
    _, status = run_git_cmd(["status", "--short"], cwd=cwd)
    files = []
    if status.strip():
        for line in status.strip().split("\n"):
            line = line.strip()
            if not line:
                continue
            parts = line.split(maxsplit=1)
            code = parts[0]
            fname = parts[1] if len(parts) > 1 else line
            
            icon = "📝"
            type_label = "تعديل"
            if "?" in code:
                icon = "✨"
                type_label = "جديد"
            elif "D" in code:
                icon = "🗑️"
                type_label = "حذف"
            elif "A" in code:
                icon = "➕"
                type_label = "مضاف"
                
            files.append({
                "code": code,
                "name": fname,
                "icon": icon,
                "type": type_label
            })
            
    return {
        "is_repo": True,
        "files": files,
        "total_changes": len(files),
        "status_clean": len(files) == 0
    }

def set_git_remote(repo_url: str = DEFAULT_PROJECT_GITHUB_URL, cwd: Optional[str] = None) -> Tuple[bool, str]:
    """تعيين أو تحديث رابط المستودع البعيد (Remote Origin)"""
    if cwd is None:
        cwd = os.path.dirname(os.path.abspath(__file__))
    clean_url = repo_url.strip() if repo_url and repo_url.strip() else DEFAULT_PROJECT_GITHUB_URL
    
    is_repo, _ = run_git_cmd(["rev-parse", "--is-inside-work-tree"], cwd=cwd)
    if not is_repo:
        run_git_cmd(["init"], cwd=cwd)
        
    s_has, _ = run_git_cmd(["remote", "get-url", "origin"], cwd=cwd)
    if s_has:
        s, out = run_git_cmd(["remote", "set-url", "origin", clean_url], cwd=cwd)
    else:
        s, out = run_git_cmd(["remote", "add", "origin", clean_url], cwd=cwd)
        
    if s:
        return (True, f"تم ربط المستودع بنجاح: {clean_url}")
    return (False, out)

def export_project_to_github(
    repo_url: str = DEFAULT_PROJECT_GITHUB_URL,
    commit_message: str = "Initial commit: Mobile Specs Dashboard",
    branch_name: str = "main",
    cwd: Optional[str] = None
) -> Tuple[bool, str]:
    """تصدير وتهيئة المشروع بالكامل ورفعه إلى GitHub"""
    if cwd is None:
        cwd = os.path.dirname(os.path.abspath(__file__))
    
    clean_url = repo_url.strip() if repo_url and repo_url.strip() else DEFAULT_PROJECT_GITHUB_URL
    branch = branch_name.strip() if branch_name and branch_name.strip() else "main"
    msg = commit_message.strip() if commit_message and commit_message.strip() else "feat: export Mobiles-spec project to GitHub"
    
    # 0. فحص أمان مسبق لمنع رفع مجلدات build و dist
    build_dirs = check_large_build_dirs(cwd=cwd)
    if build_dirs:
        names = ", ".join([f"'{d['name']}' ({d['size_mb']} MB)" for d in build_dirs])
        return (False, f"🚨 تحذير أمني: تم إيقاف الرفع لوجود فهارس ملفات EXE ضخمة ({names}). يرجى حذفها أولاً للمتابعة.")
        
    logs = []
    
    # 1. تهيئة المستودع إذا لم يكن موجوداً
    s, out = run_git_cmd(["init"], cwd=cwd)
    logs.append(f"• git init: {out if out else 'تم تهيئة المستودع بنجاح'}")
    
    # 2. تعيين اسم الفرع (Branch)
    s, out = run_git_cmd(["branch", "-M", branch], cwd=cwd)
    logs.append(f"• git branch -M {branch}: {out if out else 'تم تعيين الفرع الرئيسي'}")
    
    # 3. ضبط الرابط البعيد (Remote Origin)
    s_has, _ = run_git_cmd(["remote", "get-url", "origin"], cwd=cwd)
    if s_has:
        run_git_cmd(["remote", "set-url", "origin", clean_url], cwd=cwd)
    else:
        run_git_cmd(["remote", "add", "origin", clean_url], cwd=cwd)
    logs.append(f"• git remote set origin {clean_url}: تم ضبط المستودع البعيد")
    
    # 4. إضافة كافة الملفات (git add -A)
    s, out = run_git_cmd(["add", "-A"], cwd=cwd)
    logs.append(f"• git add -A: {out if out else 'تم تجهيز وتضمين كافة ملفات المشروع'}")
    
    # 5. عمل الـ Commit
    s_st, st_out = run_git_cmd(["status", "--short"], cwd=cwd)
    if st_out.strip():
        s, out = run_git_cmd(["commit", "-m", msg], cwd=cwd)
        logs.append(f"• git commit: {out if out else 'تم حفظ الـ Commit بنجاح'}")
    else:
        logs.append("• git commit: لا توجد تغييرات جديدة للملفات، المستودع المحلي جاهز")
    
    # 6. الرفع إلى GitHub
    s_push, push_out = run_git_cmd(["push", "-u", "origin", branch], cwd=cwd)
    logs.append(f"• git push -u origin {branch}: {push_out if push_out else 'تم الرفع بنجاح'}")
    
    if not s_push:
        # محاولة المزامنة والرفع مع allow-unrelated-histories أو force
        run_git_cmd(["pull", "origin", branch, "--rebase", "--allow-unrelated-histories"], cwd=cwd)
        s_push2, push_out2 = run_git_cmd(["push", "-u", "origin", branch], cwd=cwd)
        if not s_push2:
            s_push2, push_out2 = run_git_cmd(["push", "-u", "origin", branch, "--force"], cwd=cwd)
            logs.append(f"• git push -u origin {branch} (--force retry): {push_out2 if push_out2 else 'تم الرفع بنجاح'}")
        if s_push2:
            return (True, "\n".join(logs))
        return (False, "\n".join(logs) + "\n\n💡 تأكد من صحة الرابط وأن لديك صلاحيات الرفع (Personal Access Token / SSH Key).")
        
    return (True, "\n".join(logs))

def update_project_on_github(
    commit_message: str = "feat: update Mobiles-spec data & specs",
    branch_name: str = "main",
    repo_url: str = DEFAULT_PROJECT_GITHUB_URL,
    cwd: Optional[str] = None
) -> Tuple[bool, str]:
    """تحديث ورفع آخر التغييرات إلى GitHub وعمل Refresh"""
    if cwd is None:
        cwd = os.path.dirname(os.path.abspath(__file__))
    
    clean_url = repo_url.strip() if repo_url and repo_url.strip() else DEFAULT_PROJECT_GITHUB_URL
    branch = branch_name.strip() if branch_name and branch_name.strip() else "main"
    msg = commit_message.strip() if commit_message and commit_message.strip() else "feat: update Mobiles-spec data & specs"
    
    # 0. فحص أمان مسبق لمنع رفع مجلدات build و dist
    build_dirs = check_large_build_dirs(cwd=cwd)
    if build_dirs:
        names = ", ".join([f"'{d['name']}' ({d['size_mb']} MB)" for d in build_dirs])
        return (False, f"🚨 تحذير أمني: تم إيقاف التحديث لوجود فهارس ملفات EXE ضخمة ({names}). يرجى حذفها أولاً للمتابعة.")
        
    logs = []
    
    # التأكد من تهيئة المستودع
    is_repo, _ = run_git_cmd(["rev-parse", "--is-inside-work-tree"], cwd=cwd)
    if not is_repo:
        run_git_cmd(["init"], cwd=cwd)
        logs.append("• git init: تم تهيئة المستودع")
    
    # التأكد من اسم الفرع
    run_git_cmd(["branch", "-M", branch], cwd=cwd)
    
    # التحقق من وجود Remote Origin وضبطه
    s_rem, remote = run_git_cmd(["remote", "get-url", "origin"], cwd=cwd)
    if not s_rem or not remote or remote.startswith("error"):
        run_git_cmd(["remote", "add", "origin", clean_url], cwd=cwd)
        logs.append(f"• git remote add origin {clean_url}: تم ربط المستودع البعيد")
    
    # 1. إضافة الملفات المعدلة
    s, out = run_git_cmd(["add", "-A"], cwd=cwd)
    logs.append(f"• git add -A: {out if out else 'تم تجهيز كافة التعديلات'}")
    
    # 2. عمل Commit إذا كانت هناك تعديلات
    s_st, st_out = run_git_cmd(["status", "--short"], cwd=cwd)
    if st_out.strip():
        s, out = run_git_cmd(["commit", "-m", msg], cwd=cwd)
        logs.append(f"• git commit: {out if out else 'تم حفظ الـ Commit بنجاح'}")
    else:
        logs.append("• git commit: لا توجد تعديلات جديدة معلقة للحفظ محلياً")
    
    # 3. محاولة سحب التحديثات ثم الرفع
    run_git_cmd(["pull", "origin", branch, "--rebase", "--autostash"], cwd=cwd)
    
    s_push, push_out = run_git_cmd(["push", "-u", "origin", branch], cwd=cwd)
    logs.append(f"• git push -u origin {branch}: {push_out if push_out else 'تم الرفع والتحديث بنجاح'}")
    
    if not s_push:
        s_push2, push_out2 = run_git_cmd(["push", "-u", "origin", branch, "--force"], cwd=cwd)
        logs.append(f"• git push (force retry): {push_out2 if push_out2 else 'تم الرفع بنجاح'}")
        if s_push2:
            return (True, "\n".join(logs))
        return (False, "\n".join(logs) + "\n\n💡 يرجى التأكد من صلاحيات الحساب وربط الـ Remote بشكل صحيح.")
        
    return (True, "\n".join(logs))

def get_git_details(cwd: Optional[str] = None) -> Dict[str, Any]:
    """استرجاع تفاصيل شاملة عن المستودع والملفات والتاريخ"""
    if cwd is None:
        cwd = os.path.dirname(os.path.abspath(__file__))
    
    is_repo, _ = run_git_cmd(["rev-parse", "--is-inside-work-tree"], cwd=cwd)
    if not is_repo:
        return {
            "is_repo": False,
            "branch": "main",
            "remote": DEFAULT_PROJECT_GITHUB_URL,
            "web_url": DEFAULT_PROJECT_WEB_URL,
            "status": "المستودع المحلي غير مهيأ",
            "commits": [],
            "files": [],
            "branches": ["main"]
        }
    
    _, branch = run_git_cmd(["branch", "--show-current"], cwd=cwd)
    s_rem, remote = run_git_cmd(["remote", "get-url", "origin"], cwd=cwd)
    _, status = run_git_cmd(["status", "--short"], cwd=cwd)
    _, files_out = run_git_cmd(["ls-files"], cwd=cwd)
    _, log_out = run_git_cmd(["log", "-15", "--pretty=format:%h|%an|%ar|%s"], cwd=cwd)
    _, branches_out = run_git_cmd(["branch", "-a"], cwd=cwd)
    
    clean_remote = remote if (s_rem and remote and not remote.startswith("error")) else DEFAULT_PROJECT_GITHUB_URL
    
    commits_list = []
    if log_out:
        for line in log_out.strip().split("\n"):
            parts = line.split("|", 3)
            if len(parts) == 4:
                commits_list.append({
                    "hash": parts[0],
                    "author": parts[1],
                    "date": parts[2],
                    "msg": parts[3]
                })
                
    files_list = [f for f in files_out.strip().split("\n") if f] if files_out else []
    
    clean_web_url = DEFAULT_PROJECT_WEB_URL
    if clean_remote != "-" and "github.com" in clean_remote:
        clean_web_url = clean_remote.replace(".git", "").replace("git@github.com:", "https://github.com/").strip()
        
    return {
        "is_repo": True,
        "branch": branch if branch else "main",
        "remote": clean_remote,
        "web_url": clean_web_url,
        "status": status if status else "المستودع محدث ونظيف",
        "commits": commits_list,
        "files": files_list,
        "branches": [b.strip() for b in branches_out.strip().split("\n") if b] if branches_out else []
    }


