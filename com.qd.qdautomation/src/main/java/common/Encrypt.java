package common;



import org.apache.commons.codec.binary.Base64;

import javax.crypto.Cipher;
import java.security.KeyFactory;
import java.security.KeyStore;
import java.security.interfaces.RSAPrivateKey;
import java.security.interfaces.RSAPublicKey;
import java.security.spec.PKCS8EncodedKeySpec;
import java.security.spec.X509EncodedKeySpec;

public class Encrypt {
    public KeyStore kestore = null;
    public KeyStore enkestore = null;
    public Cipher chiher = null;
    public Cipher enchiher = null;

    public Encrypt() {
        try {
            String pubKey = "MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA6grSvCTxQnvWn7X2R4LCHXDY8KZARn7c4GhndIiQK95yQiLqgrt0m9HD0qlCcHEYZDyAdELITaac/EmN4MlbAFzOCy2vZrxmOTYXXK/WhufzWL/4RP8mmxUMGAMQ+mYnfv6shHUjmy0MAYq/2jb3WHcA9rhc84E8u+et0KA/p2urWaMX5glSqo9LYl87b7pvbIv9D0PKX+6u4aMRzwcZ3pe0tCvGx0lPkzznqPl2MMq0grpG+2+R35KzihyVjuNJYzuaWhxnHfRHcMrE4sAl8gh48FHAAOcBtccn2ElOWDkCXpPgPCYQ/dlmz3gLGJVaztAvWZvhRTs3cdj1DoXxWQIDAQAB";
            String privateKey = "MIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQDqCtK8JPFCe9aftfZHgsIdcNjw\r\npkBGftzgaGd0iJAr3nJCIuqCu3Sb0cPSqUJwcRhkPIB0QshNppz8SY3gyVsAXM4LLa9mvGY5Nhdc\r\nr9aG5/NYv/hE/yabFQwYAxD6Zid+/qyEdSObLQwBir/aNvdYdwD2uFzzgTy7563QoD+na6tZoxfm\r\nCVKqj0tiXztvum9si/0PQ8pf7q7hoxHPBxnel7S0K8bHSU+TPOeo+XYwyrSCukb7b5HfkrOKHJWO\r\n40ljO5paHGcd9EdwysTiwCXyCHjwUcAA5wG1xyfYSU5YOQJek+A8JhD92WbPeAsYlVrO0C9Zm+FF\r\nOzdx2PUOhfFZAgMBAAECggEAaSYbr7PZQMlqj66sy0wP8jI1BUlsLz/n9FpzwINqYos6QkR1n+6g\r\nYTFJDWhGFUZ6cu5Lt/AZFen5Nc/GEmvkmrEXUSZhW4LhWmrW7CmD9HBZgFIfNMB9k8dM+6zE//rN\r\nW90HnSsTqDAA6OBxeZSgY/pzUGVcPXqe0r0v9j9I2epFvqdsgtHr0APtboPAcM0bqK4Ba9kE5nMo\r\np2hA7epORL3vsQZnHyfdcmHukzhtrVsLIOqWWEQxDP/cZnWLFOqasqCGhYqg7kYDJHvuW8ioF3co\r\nP3STggjbPbJY8ZIapZmrer5/JRWQHKbru/cXrIQ2ys1umsiL3XGx00oedU4IAQKBgQD6HS3Qs8T/\r\nouLn3r43dppiFtNz6TdtAp2nAwqa37CIfAxPOpZdfV0opkJRNBus+1K+T+cZjo1OVxSimboJzkpa\r\niTfdfzy7lebETPYJPLPLLiYYW7E1rRx7Qzx0s3qy/6yguU7TfhTEbkl+uuEiImbRgIrd+RxFGpkQ\r\nonP1uzSewQKBgQDvjNHUTjUIsVtg4nU9P1zxtKZl7/KNkj95x5CRLGFHJHbWwWqiGi5mEBM8pmYS\r\nnAW2T9l6ibFR4tZjy2YvJvgE490OAjKPHC5UjPj7S2z39cwNZ2igY3kHjjYP7o/UsnsxF5FEN4F7\r\nbA/0k0unX7vPZEiBVXJNReK9kE7UIucQmQKBgFcGXhGW6z2TXGYXP5Imx1LH5G0ZQXrBhT8+NKts\r\nl4z1tIbAuN7wpsBlEQAWJGcMz15+iMeArTQL4pcs6+lLiN11jyWXhvcyEihNKvJN17UjaofhJ4io\r\nohdVbqygYlnz0gneUEoQPYOXaCUbOwhYdfQBSNAxRB/bQFTkYXqn14hBAoGBAN0WdNqjS6M4PYQP\r\nvfDHODsEZccTTCOJZPg6TZ0O/sEwUFYosQfnbgdYssh5Mx/dzA/VZn0H9BYFa6UzJp5CMwogW/b/\r\n5DXc5o22hl034dyAiNGeLRxKcnucrxzJKwo9qCFJBROWvvRwadFq2Jojay0f+yMS/6l95KTcOUTk\r\n6amRAoGAKNMQxK+VNAXRAZ9gX5Y9dniXrAvdn3jS3CjBQzRo/ET6okhGQXA7bmVMrAGGY9DA+uue\r\nmIP0WZ9D2bh1k0rCjG5Eb89LjFsQ7L579pgE3MooCLum9PNvQL/Z8rLLF66Y5pwIKJhXLrkMiU0w\r\nd8MwljJ6LeXlb1Fir5dE2M9V7RA=";
            KeyFactory keyFactory = KeyFactory.getInstance("RSA");
            X509EncodedKeySpec x509KeySpec = new X509EncodedKeySpec(Base64.decodeBase64(pubKey));
            RSAPublicKey key = (RSAPublicKey)keyFactory.generatePublic(x509KeySpec);
            PKCS8EncodedKeySpec pkcs8KeySpec = new PKCS8EncodedKeySpec(Base64.decodeBase64(privateKey));
            RSAPrivateKey pkey = (RSAPrivateKey)keyFactory.generatePrivate(pkcs8KeySpec);
            this.enchiher = Cipher.getInstance("RSA");
            this.enchiher.init(1, key);
            this.chiher = Cipher.getInstance("RSA");
            this.chiher.init(2, pkey);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    public String deCrypt(String en) {
        byte[] bs = null;
        try {
            if (en != null && en.length() > 300) {
                bs = Base64.decodeBase64(en);
                return new String(this.chiher.doFinal(bs));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return null;
    }

    public String enCrypt(String en) {
        try {
            if (en != null && en.length() > 0) {
                byte[] doFinal = this.enchiher.doFinal(en.getBytes());
                return this.encryptBASE64(doFinal);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        return null;
    }

    private String encryptBASE64(byte[] key) throws Exception {
        return Base64.encodeBase64String(key).replace("\r", "").replace("\n", "");
    }
}
