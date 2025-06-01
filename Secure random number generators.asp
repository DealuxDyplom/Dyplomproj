using Org.BouncyCastle.Security;

byte[] random = new byte[8];

SecureRandom sr = SecureRandom.GetInstance("SHA256PRNG", false);
sr.NextBytes(random); // Noncompliant

using Org.BouncyCastle.Crypto.Digest;
using Org.BouncyCastle.Crypto.Prng;

byte[] random = new byte[8];

IRandomGenerator digest = new DigestRandomGenerator(new Sha256Digest());
digest.AddSeedMaterial(Encoding.UTF8.GetBytes("predictable seed value"));
digest.NextBytes(random); // Noncompliant

IRandomGenerator vmpc = new VmpcRandomGenerator();
vmpc.AddSeedMaterial(Convert.FromBase64String("2hq9pkyqLQJkrYTrEv1eNw=="));
vmpc.NextBytes(random); // Noncompliant

using Org.BouncyCastle.Crypto.Digest;
using Org.BouncyCastle.Crypto.Prng;
using Org.BouncyCastle.Security;

byte[] random = new byte[8];

IRandomGenerator digest = new DigestRandomGenerator(new Sha256Digest());
SecureRandom sr = new SecureRandom(digest);
sr.NextBytes(random); // Noncompliant