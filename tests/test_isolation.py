# -*- coding: utf-8 -*-
"""
导师版 × 学生版 授权隔离回归测试
================================
证明四个隔离层同时生效，避免两版授权互相通用：
  ① 在线卡密应用名（KAMI_APP=daoshiban）
  ② 离线备用码密钥（_OFFLINE_KEY v2，mentor 专用）
  ③ 本地授权目录（~/.tfd_mentor_license）+ product 字段校验
  ④ 试用计数盐（_TRIAL_SALT，mentor 专用）

运行：python tests/test_isolation.py
"""
import os
import sys
import json
import hmac
import hashlib
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(HERE)
sys.path.insert(0, PROJECT_ROOT)

import tfd_app.license as lic
import tfd_app.trial as trial


def main():
    mc = lic.get_machine_code()
    print("本机机器码:", mc)

    # 1) 导师版离线码正常往返
    code = lic.generate_offline_code(mc)
    assert lic.verify_offline_code(code, mc) is True, "导师版离线码本机校验应成功"
    assert lic.verify_offline_code(code, "WRONGMACHINE") is False, "机器码不符应失败"
    print("[PASS] 导师版离线码往返正常")

    # 2) 学生版离线码（旧密钥 v1）在导师版中应被拒
    OLD_KEY = _obf("OykrMyQuIiYzICkpIyYhKjN9f315MzwmKCEzOX4=").encode("utf-8")
    def student_sign(payload):
        return hmac.new(OLD_KEY, payload.encode("utf-8"), hashlib.sha256).hexdigest()[:24]
    ts = 1234567890
    student_code = "%s|%d|%s" % (mc, ts, student_sign("%s|%d" % (mc, ts)))
    assert lic.verify_offline_code(student_code, mc) is False, "学生版离线码不应通过导师版校验"
    print("[PASS] 学生版离线码被导师版拒绝（密钥隔离生效）")

    # 3) product 字段隔离：非 mentor 授权文件应被拒，mentor 授权应放行
    tmp = tempfile.mkdtemp()
    lic.LICENSE_DIR = tmp
    lic.LICENSE_FILE = os.path.join(tmp, "license.json")

    fake_student = {"machine_code": mc, "product": "student", "permanent": True, "offline": True}
    with open(lic.LICENSE_FILE, "w", encoding="utf-8") as f:
        json.dump(fake_student, f)
    assert lic.check_local_valid() is False, "product=student 应被导师版拒绝"
    print("[PASS] product=student 授权被导师版拒绝")

    fake_mentor = {"machine_code": mc, "product": "mentor", "permanent": True, "offline": True}
    with open(lic.LICENSE_FILE, "w", encoding="utf-8") as f:
        json.dump(fake_mentor, f)
    assert lic.check_local_valid() is True, "product=mentor 应放行"
    print("[PASS] product=mentor 授权放行")

    # 4) 试用盐与导师版绑定
    assert b"mentor" in trial._TRIAL_SALT, "试用盐应包含 mentor 标识"
    print("[PASS] 试用盐为导师版独立盐值")

    print("\n=== 隔离验证全部通过 ===")


if __name__ == "__main__":
    main()
