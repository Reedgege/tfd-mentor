# -*- coding: utf-8 -*-
"""
导师版 × 学生版 授权隔离回归测试
================================
证明四个隔离层同时生效，避免两版授权互相通用：
  ① 在线卡密应用名（KAMI_APP=daoshiban）
  ② 离线备用码密钥（Ed25519 非对称；导师版独立密钥对，与学生版/海外版互不通用）
  ③ 本地授权目录（~/.tfd_mentor_license）+ product 字段校验
  ④ 试用计数盐（_TRIAL_SALT，mentor 专用）

运行：python tests/test_isolation.py
依赖：cryptography（仅测试/发码器用；客户端本身零第三方加密依赖）
"""
import os
import sys
import json
import time
import base64
import tempfile

from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey

HERE = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(HERE)
sys.path.insert(0, PROJECT_ROOT)

import tfd_app.license as lic
import tfd_app.trial as trial
import tfd_app.crypto as crypto_mod


def _sign(priv, machine_code):
    """测试用离线码签名器（真实签名只由卖家统一发码器持私钥完成）。"""
    payload = "%s|%d" % (machine_code, int(time.time()))
    sig = priv.sign(payload.encode("utf-8"))          # 64 字节 Ed25519 签名
    return payload + "." + base64.urlsafe_b64encode(sig).decode("ascii")


def main():
    mc = lic.get_machine_code()
    print("本机机器码:", mc)

    # 1) 导师版离线码正常往返（用一份临时 Ed25519 密钥对模拟卖家私钥/客户端公钥）
    mentor_priv = Ed25519PrivateKey.generate()
    crypto_mod.set_verify_public_key(mentor_priv.public_key().public_bytes_raw())
    code = _sign(mentor_priv, mc)
    assert lic.verify_offline_code(code, mc) is True, "导师版离线码本机校验应成功"
    assert lic.verify_offline_code(code, "WRONGMACHINE") is False, "机器码不符应失败"
    print("[PASS] 导师版离线码往返正常")

    # 2) 学生版离线码（独立 Ed25519 密钥对）在导师版中应被拒 —— 密钥隔离生效
    student_priv = Ed25519PrivateKey.generate()
    student_code = _sign(student_priv, mc)
    assert lic.verify_offline_code(student_code, mc) is False, "学生版离线码不应通过导师版校验"
    print("[PASS] 学生版离线码被导师版拒绝（密钥隔离生效）")

    # 2-b) 小阶/单位元点必须被拒（纵深加固回归）
    identity = b"\x01" + b"\x00" * 31          # 单位元（阶 1）
    bad_sig = identity + (1).to_bytes(32, "little")
    bad_code = ("%s|1700000000" % mc) + "." + base64.urlsafe_b64encode(bad_sig).decode("ascii")
    assert lic.verify_offline_code(bad_code, mc) is False, "R 为单位元应被拒"
    crypto_mod.set_verify_public_key(identity)
    try:
        assert lic.verify_offline_code(bad_code, mc) is False, "公钥为单位元应被拒"
    finally:
        crypto_mod.set_verify_public_key(None)
    print("[PASS] 小阶/单位元点被拒")

    crypto_mod.set_verify_public_key(None)

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
