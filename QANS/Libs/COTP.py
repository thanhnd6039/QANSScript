import pyotp
class COTP(object):
    def generate_otp(self, key):
        totp = pyotp.TOTP(key)
        otp = totp.now()
        return otp

