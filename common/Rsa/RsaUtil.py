#!/usr/bin/env python
# -*- coding: utf-8 -*-

import base64
from m2crypto.M2Crypto import RSA
class RsaUtil:

    @staticmethod
    def encrypt(digest, private_key):
        digest=base64.b64encode(digest)
        result = ""
        while (len(digest) > 117):
            some = digest[0:117]
            digest = digest[117:]
            result += private_key.private_encrypt(some, RSA.pkcs1_padding).encode("hex")

        result += private_key.private_encrypt(digest, RSA.pkcs1_padding).encode("hex")

        return result

if __name__ == "__main__":
    private_key = RSA.load_key('8000013189_pri.pem')
    result= RsaUtil.encrypt("123456",private_key)
    print(result)
