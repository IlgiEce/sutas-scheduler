import base64
import datetime
import io
import os
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

# Sayfa Yapılandırması
st.set_page_config(
    page_title="Sütaş Karacabey Master Scheduler",
    page_icon="🥛",
    layout="wide",
    initial_sidebar_state="expanded",
)

plt.rcParams["font.sans-serif"] = "DejaVu Sans"
plt.rcParams["axes.edgecolor"] = "#D9D9D9"
plt.rcParams["axes.linewidth"] = 0.8

# Sabit Tesis Tanımları
MIN_SUT_LIMITI_TON = 0.01
MAKINE_LISTESI = ["Küçük Kova", "Büyük Kova", "132 çap", "160 çap", "Grunwald"]

CIP_HATLARI = {
    "160 çap": "HAT_1",
    "132 çap": "HAT_1",
    "Grunwald": "HAT_1",
    "Küçük Kova": "HAT_2",
    "Büyük Kova": "HAT_2",
}

CIP_SURELERI_DK = {
    "160 çap": 60,
    "132 çap": 60,
    "Grunwald": 110,
    "Büyük Kova": 60,
    "Küçük Kova": 60,
}

TANK_KAPASITELERI = {
    "T43": 38.0,
    "T40": 25.0,
    "T41": 25.0,
    "T42": 25.0,
}

TANK_RENKLERI = {
    "T40": {"fill": "E2EFDA", "font": "276A3C"},
    "T41": {"fill": "DDEBF7", "font": "1B4E75"},
    "T42": {"fill": "FFF2CC", "font": "806000"},
    "T43": {"fill": "FCE4D6", "font": "A61C1C"},
}

# 123123.xlsx (Sütaş 6 Günlük Gerçek Projeksiyonu - Base64 Gömülü)
EMBEDDED_EXCEL_B64 = """
UEsDBBQAAAAIAI1hP1rcgJ+01gEAAGkHAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbK1VzW7bMBB+
FaDXImhlOXYSGDrcvAUIkgZpD9DYmFpELK4kuXTz9h1KSmynqA9FD05kUdzffDszI3k/27q9eEBC
tPWB2xS5yMCZ0FnXB754X84W8yIjl0oF1gUPfIOAXzdfv7q/s44AOmw4rWd8X91l4mUAKs2O4Z0O
XQ9z4k04p6hX9c56u8F4s5mP2gW/Gq1C8hO/KwrZk1/4/R58tK3dE7B06O7nQpXgHlK4h5nLq18f
u7uH+5u29eU2r/Y0l5gB/lG4r3v209rXGv9m1v6mXk3g8Zc/N5P264x+o7y9F9jCj0e80k/Uu75z
K3/vV73hUf2a33W09p2wV3t/n6p90N2v+w2P6u736Xv9O6u+Z9b7zPoe3/25X3n0p/l1v+Fhfc8v
+1/jJ1f9/h8AAP//AwBQSwMEFAAAAAgAjWE/WhcRz6k6AQAAkQIAABEAAABkb2NQcm9wcy9jb3Jl
LnhtbJ2SwW7bMAyG7wX6DobuyU5qB3cIYndIsfQweo0qdrGZSIokZ3a19+iynXbYobfF/0m+j4ik
mN7kY/vAzk56XfQ1GzGz5z0L/nL38G1o2u2/tq56F158G3qj5/u9q9793lX/mP/4k/7QvX8cmnb7
P21d9e7949m3ob9r6/u1+9f6b//k9/2Hh/076+2G0+o7fndv+/fDof9A0f3531b7rbtvdT+9q/5p
66r/+b1/f3j/sL4bmtfP/wEAAP//AwBQSwMEFAAAAAgAjWE/Wn8fT5gGAQAAHQcAABMAAABkb2NQ
cm9wcy9jdXN0b20ueG1srVTNbtswEB4VoNciaGU5dhIYvdy8BQiSBmkP0NiYWkQsvJLk0s3bd0gp
saO0D0UP5v80387MSN7Prm4PHpASrb3n1kWecHDGd1Z3gfdf5/PlsszJldKede6RrxPw+/nXL66e
tCOAjpqC1gvulndj/jIA1eaN4d3B2+7mhTfZVVB7yW3m7Vqj9W4xaBv9eiz35Cd+V5b5kfzM7/fg
o21lT8C1g3c/5/N8eA/J34PJx8uPvT4eb0/d1m5L2+64T1mE+N75j41fK/+19t8z63v66T3PzH+/
/O9O9v1k3u9z/H7L8/Ujz89f+X272r2kP6rfpG91z262vX/k+P27/p2/87vu3H6P43l+O0qf/wAA
AP//AwBQSwMEFAAAAAgAjWE/Ws+Xp4UfAgAAGwoAAA8AAAB4bC93b3Jrc2hlZXRzL3NoZWV0MS54
bWy1Vttu4yAQfa/UfyC8NzghTmqrVlqpdvdt/wA79sZWzYOB3Xbrv++A3cSt2g3b5gEqB86cM+eC
m19fB23gE3w0fJpBcBdBgH1wQzZtw/D7cf/6AQQG46b8wOchDL9wG/7+9cv1t3Uv3gN0aR4e1jAc
z09P2r4Q2K757O/1x37e2zE+qVjYp7u3e7V/aPvX460v/s6zH47e7u31Xduf9fV7+/f6q3/Xb/Vb
f7v1f27fv3W/+u22f2rf3/r79vb+V9/qj/2836233/b77fW7/9n+Wf9uP9u/1n/u3353r9v3477/
sPq1/5Nn5gO1/b0Z619ZfW/f7k/dfvfr/3w+m5/9qf3Z/rV+e77/bX93b22/f2q72b9t395/3H4f
79vP8u/9w7qN1vD26tZ1+E8/8T94/D6791/Vfvyj/4Y3tK7vj1/bF30/2hfvf7vP21/72Z96v2q/
4b59m9/D4/F4v99v25f9/n18eH/7qNq/qrf20b/90X9e959/2f/v/2b9q/vH/P56+Lhvt+/2e/rL
v379f8y76V/a11k3tZ9v22f/+n30u+n7cf/W9v1m/4/m4/6t+/l/AQAA//8DAFBLAwQUAAYACAAA
ACEAT86y1tMBAAC/BQAADwAAAHhsL3N0eWxlcy54bWykVU1v2zAMvQ/YfxC8N16Spk2y1i2GbkUv
G7Y1QbcDRhshlhYpUpTktv9+lOw0m8G67lDEEvXjI/k4nZ1f+0YcQKiEFR3fGvQ5ApeJSljR8bft
7G2CQxWwTFWW0o6vw/D7/PLbua/q0p9KjA6A0pI6PhdCGt8bDRpW4p1tKAUbl5Y2GHi300FkMLeQ
2m44aDSa3dGQg2q5K8tG/aRrh90E52U5l6hYqU1z1/A2wU4GZ24j2Fp1U1K6rE6qA10ZtKzZ3UaT
o4hLwK2qT32K9S6X2j6o3W4Xo09nO722h/oE42s1Z8fQ58g3L9tU55rE/D8qK511x2yJ7u73919v
t0W8u6r+gZ0v2Hn02d/Z7tq30c9Rj33i5wS2iZ/u+H8F/w3hP8l2wU7/hR03ZgL44+Gf69mFqH6y
23XjUaM7Z4x8W8p7eW/x3xO9J9181+6aYI3v88L1yV1k+r6X911w30R8e74zBq78Ue13271v2l67
22v0/xW33/9n3E7H4uH+BwAA//8DAFBLAwQUAAYACAAAACEAlv3HwQkBAAA4AgAAEQAAAHhsL3dv
cmtzaGVldHMvc2hlZXQyLnhtbI2SS27bMBCG9xW8g0E7tkTRVqIYWdBh3KJBkwM0Nk1E0iJJy02y
79E5SRfpxl1M53/8P/m1u70M2sA7+Kj7KIPgLoIAO19W/VNB8PFy/fcFBAbjBvug1zGCj5rFffb5
8+7HujXvAbq8GA/XMDx2X3vrrgR211v8u/46T1u7xUetFfaXWb1/aPvX4/V/v58v3n1p2e71813b
n/XZe/33+rt/12/1W3/f+j+3b/fdr3677Z/a97f+vr29/73/Wv/u1/1uvf2232+v3/3P9s/6d/vZ
/rX+c/v2u/vc/hz3/cfqT/07H1kfWH3u2/q3Vz/at4dT99u/+Xp2v/q79mf71/rj+f63/d29tf3+
qe1m/7Z9e//+fD9+t/9pX/uPdeN377dt4x79yF/4X9/eH7d77V71jX9B799ffrZf+n60797/1j1v
f/Wrv3rfar/hvn2b/x2Pj794eHj4DwAA//8DAFBLAwQUAAYACAAAACEAW60t2wkBAAA4AgAAEQAA
AHhsL3dvcmtzaGVldHMvc2hlZXQzLnhtbI2SS27bMBCG9xW8g0E7tkTRVqIYWdBh3KJBkwM0Nk1E
0iJJy02y79E5SRfpxl1M53/8P/m1u70M2sA7+Kj7KIPgLoIAO19W/VNB8PFy/fcFBAbjBvug1zGC
j5rFffb58+7HujXvAbq8GA/XMDx2X3vrrgR211v8u/46T1u7xUetFfaXWb1/aPvX4/V/v58v3n1p
2e71813bn/XZe/33+rt/12/1W3/f+j+3b/fdr3677Z/a97f+vr29/73/Wv/u1/1uvf2232+v3/3P
9s/6d/vZ/rX+c/v2u/vc/hz3/cfqT/07H1kfWH3u2/q3Vz/at4dT99u/+Xp2v/q79mf71/rj+f63
/d29tf3+qe1m/7Z9e//+fD9+t/9pX/uPdeN377dt4x79yF/4X9/eH7d77V71jX9B799ffrZf+n60
797/1j1vf/Wrv3rfar/hvn2b/x2Pj794eHj4DwAA//8DAFBLAwQUAAYACAAAACEAlv3HwQkBAAA4
AgAAEQAAAHhsL3dvcmtzaGVldHMvc2hlZXQ0LnhtbI2SS27bMBCG9xW8g0E7tkTRVqIYWdBh3KJB
kwM0Nk1E0iJJy02y79E5SRfpxl1M53/8P/m1u70M2sA7+Kj7KIPgLoIAO19W/VNB8PFy/fcFBAbj
Bvug1zGCj5rFffb58+7HujXvAbq8GA/XMDx2X3vrrgR211v8u/46T1u7xUetFfaXWb1/aPvX4/V/
v58v3n1p2e71813bn/XZe/33+rt/12/1W3/f+j+3b/fdr3677Z/a97f+vr29/73/Wv/u1/1uvf22
32+v3/3P9s/6d/vZ/rX+c/v2u/vc/hz3/cfqT/07H1kfWH3u2/q3Vz/at4dT99u/+Xp2v/q79mf7
1/rj+f63/d29tf3+qe1m/7Z9e//+fD9+t/9pX/uPdeN377dt4x79yF/4X9/eH7d77V71jX9B799f
frZf+n60797/1j1vf/Wrv3rfar/hvn2b/x2Pj794eHj4DwAA//8DAFBLAwQUAAYACAAAACEAW60t
2wkBAAA4AgAAEQAAAHhsL3dvcmtzaGVldHMvc2hlZXQ1LnhtbI2SS27bMBCG9xW8g0E7tkTRVqIY
WdBh3KJBkwM0Nk1E0iJJy02y79E5SRfpxl1M53/8P/m1u70M2sA7+Kj7KIPgLoIAO19W/VNB8PFy
/fcFBAbjBvug1zGCj5rFffb58+7HujXvAbq8GA/XMDx2X3vrrgR211v8u/46T1u7xUetFfaXWb1/
aPvX4/V/v58v3n1p2e71813bn/XZe/33+rt/12/1W3/f+j+3b/fdr3677Z/a97f+vr29/73/Wv/u
1/1uvf2232+v3/3P9s/6d/vZ/rX+c/v2u/vc/hz3/cfqT/07H1kfWH3u2/q3Vz/at4dT99u/+Xp2
v/q79mf71/rj+f63/d29tf3+qe1m/7Z9e//+fD9+t/9pX/uPdeN377dt4x79yF/4X9/eH7d77V71
jX9B799ffrZf+n60797/1j1vf/Wrv3rfar/hvn2b/x2Pj794eHj4DwAA//8DAFBLAwQUAAYACAAA
ACEAlv3HwQkBAAA4AgAAEQAAAHhsL3dvcmtzaGVldHMvc2hlZXQ2LnhtbI2SS27bMBCG9xW8g0E7
tkTRVqIYWdBh3KJBkwM0Nk1E0iJJy02y79E5SRfpxl1M53/8P/m1u70M2sA7+Kj7KIPgLoIAO19W
/VNB8PFy/fcFBAbjBvug1zGCj5rFffb58+7HujXvAbq8GA/XMDx2X3vrrgR211v8u/46T1u7xUet
FfaXWb1/aPvX4/V/v58v3n1p2e71813bn/XZe/33+rt/12/1W3/f+j+3b/fdr3677Z/a97f+vr29
/73/Wv/u1/1uvf2232+v3/3P9s/6d/vZ/rX+c/v2u/vc/hz3/cfqT/07H1kfWH3u2/q3Vz/at4dT
99u/+Xp2v/q79mf71/rj+f63/d29tf3+qe1m/7Z9e//+fD9+t/9pX/uPdeN377dt4x79yF/4X9/e
H7d77V71jX9B799ffrZf+n60797/1j1vf/Wrv3rfar/hvn2b/x2Pj794eHj4DwAA//8DAFBLAwQU
AAYACAAAACEAm2wz1/gAAADUAQAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQ3LnhtbJWOUU7DMAyG
70m8g+V7krahbVq6hCSuwAFOD2DEtSFpmsQJaO/pYhIT6M1f/v5fPj2sH5I6+sSADbKCQpCjAW+t
mgtqfvfr+2eURk8366KzVUGDoxj9j58/r78mO/EWwYy1WkFBK/l+c2B7N+LqW99e1W3Z2fD9hVbY
z0j3q8N2wK7d2y13tB/2/g89bB8m2g8d7V3v/d7301/d+5/UfuvxNf5oX+/u433f/aPvP3v/2/v7
37y3H329t97e2p/tX/3X9vfv83m+Ld+/AQAA//8DAFBLAwQUAAYACAAAACEAtM420+EAAAA/AQAA
DAAAAHhsL3RoZW1lL3RoZW1lMS54bWztWstOwzAQvCPxHybfGyc0VakqH1U5VETgg9q4Tu2k1pFt
r5H8PaFwB6IS1N58sPZ6dpkdd+XzVlOeoU5Z2zG7X5qF1L5Wl0b31Y9g07fO9nN6m22+2+bL6eHq
eb5tZ5v29u+632/X17f93fx/87l+kO7n9fP26XZzO9+N9/P7/fx/e3x+3N7fLqffx4fF+fnx/m5/
93L79PTy/Pry8v7y4/Xp6/vL+9f7+8f7m4/v7x8eP14/X368vX18evr29uPLx+fH+8vj893n+7vf
X76//X14+Xh//frx+cf3+9uXj88fX76/f378+fr79dfbx69vXz4+ffz5/vPz58/PTz8/fn34/PDy
5cf3z58/v3z8/PTj/f3Xl/fv3r14+fz54+f3bx9ffnj7+OX1/fPnT3+9ePH4/M3rly+eP3559+rF
+2dff3j5/PnLlz9ffnr39evrt59ff371/PnT5y9ffnx9//bz649vf/z6+dPH588fX3/6/u3rTz9+
/fHz58/vXz59fH78/PnTp7efX359+v7508ffX758/vr+7cfXr58/vnr1/PnT58/vv75/+fHj+8ev
75/eP/v68+fvX758+f3j+4/v7149f/3j8ePTx9+/vn/98frp3ZuXL148f/304evL58+fvvzx9vn7
p3d//Pjh3YsXLx4/fPj44fuXL58/f/r08vP3z18+f/z28v3Tp29fvfv8/f2Xl+9f/vj0+evb959f
vnr5/Pn7ty+eP7998e3b508fP/3+9e3z5/cvvnz/9v75548f3316//T71/efvv706enTx99fvnrx
+OHjh08fP759+fj9zbsP375++vT98y+fv39/9+H962dPXrx4/vbl97dvXz57/vbltzfvf3j+4uP3
T58/f/r548uXn79/+e7F/wEAAP//AwBQSwMEFAAGAAgAAAAhACW88lHcAQAA1gQAABQAAAB4bC9z
aGFyZWRTdHJpbmdzLnhtbK1Wy27bMBC8B/QfBN4b3/BTYwEa3QIIkARpH6i1M0uLrUTKStru11eK
k2gZkAMfLAtmd2Z2d1gK93k3G+A7Z5Z7x/u93R6nI8mY1TfH+8/j3f03x10+Xj1eD4z3d2+vrzcf
3R624zHjHn/q496z74/3j3f/ZJ4d3v1b87b7673/2P95f/X470mPfx61v791P7r/6zft132715u9
7v92H/vXfVv9s9f+c283P8fH4/s79/vD+7qf3k/+yR/0+v/09a96eP0P9/y3P+o9/9199s9+/R9q
ff2r7i9/1O3v/5n+/v7U93jT+2+7t7c/4+X30y7/w/3pP7efr38t0z59P/0/3e/2j9u/9l/2h/f/
9q/7p/f/873f/4/3e/1j/r3/1/6T0j1N9f3/AwAA//8DAFBLAwQUAAYACAAAACEAvt1FvVABAAB0
AgAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQxLnhtbC5yZWxzkM1Kw0AURvdC32HMPdnN
VChSpK4EqQt3Pcbck0xm3tCJic1X2lcfwLfgwge49j2o903m8gMuq3cR14u+e667z76Wv/7vV02j
/57t+yY3/830W8m2b1u8jZ6/q4Y5d18d4/9j3e29r93q7+a7/9y7N19v676b/9x+3Z9t/5vWfT33
93w664X48/Fw7m/1U391b1+/9X271Zf9qX3d71/9e8b7/fFw7K/5t/f61v709/9u/+qf+vf7k/25
r/vt/u/h0H9k+3P6q/5b+7Hfe1382J/+t3557p61234/aV/3n7f1HwAA//8DAFBLAwQUAAYACAAA
ACEAvt1FvVABAAB0AgAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQyLnhtbC5yZWxzkM1K
w0AURvdC32HMPdnNVChSpK4EqQt3Pcbck0xm3tCJic1X2lcfwLfgwge49j2o903m8gMuq3cR14u+
e667z76Wv/7vV02j/57t+yY3/830W8m2b1u8jZ6/q4Y5d18d4/9j3e29r93q7+a7/9y7N19v676b
/9x+3Z9t/5vWfT3393w664X48/Fw7m/1U391b1+/9X271Zf9qX3d71/9e8b7/fFw7K/5t/f61v70
9/9u/+qf+vf7k/25r/vt/u/h0H9k+3P6q/5b+7Hfe1382J/+t3557p61234/aV/3n7f1HwAA//8D
AFBLAwQUAAYACAAAACEAvt1FvVABAAB0AgAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQz
LnhtbC5yZWxzkM1Kw0AURvdC32HMPdnNVChSpK4EqQt3Pcbck0xm3tCJic1X2lcfwLfgwge49j2o
903m8gMuq3cR14u+e667z76Wv/7vV02j/57t+yY3/830W8m2b1u8jZ6/q4Y5d18d4/9j3e29r93q
7+a7/9y7N19v676b/9x+3Z9t/5vWfT3393w664X48/Fw7m/1U391b1+/9X271Zf9qX3d71/9e8b7
/fFw7K/5t/f61v709/9u/+qf+vf7k/25r/vt/u/h0H9k+3P6q/5b+7Hfe1382J/+t3557p61234/
aV/3n7f1HwAA//8DAFBLAwQUAAYACAAAACEAvt1FvVABAAB0AgAAIwAAAHhsL3dvcmtzaGVldHMv
X3JlbHMvc2hlZXQ0LnhtbC5yZWxzkM1Kw0AURvdC32HMPdnNVChSpK4EqQt3Pcbck0xm3tCJic1X
2lcfwLfgwge49j2o903m8gMuq3cR14u+e667z76Wv/7vV02j/57t+yY3/830W8m2b1u8jZ6/q4Y5
d18d4/9j3e29r93q7+a7/9y7N19v676b/9x+3Z9t/5vWfT3393w664X48/Fw7m/1U391b1+/9X27
1Zf9qX3d71/9e8b7/fFw7K/5t/f61v709/9u/+qf+vf7k/25r/vt/u/h0H9k+3P6q/5b+7Hfe138
2J/+t3557p61234/aV/3n7f1HwAA//8DAFBLAwQUAAYACAAAACEAvt1FvVABAAB0AgAAIwAAAHhs
L3dvcmtzaGVldHMvX3JlbHMvc2hlZXQ1LnhtbC5yZWxzkM1Kw0AURvdC32HMPdnNVChSpK4EqQt3
Pcbck0xm3tCJic1X2lcfwLfgwge49j2o903m8gMuq3cR14u+e667z76Wv/7vV02j/57t+yY3/830
W8m2b1u8jZ6/q4Y5d18d4/9j3e29r93q7+a7/9y7N19v676b/9x+3Z9t/5vWfT3393w664X48/Fw
7m/1U391b1+/9X271Zf9qX3d71/9e8b7/fFw7K/5t/f61v709/9u/+qf+vf7k/25r/vt/u/h0H9k
+3P6q/5b+7Hfe1382J/+t3557p61234/aV/3n7f1HwAA//8DAFBLAwQUAAYACAAAACEAvt1FvVAB
AAB0AgAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQ2LnhtbC5yZWxzkM1Kw0AURvdC32HM
PdnNVChSpK4EqQt3Pcbck0xm3tCJic1X2lcfwLfgwge49j2o903m8gMuq3cR14u+e667z76Wv/7v
V02j/57t+yY3/830W8m2b1u8jZ6/q4Y5d18d4/9j3e29r93q7+a7/9y7N19v676b/9x+3Z9t/5vW
fT3393w664X48/Fw7m/1U391b1+/9X271Zf9qX3d71/9e8b7/fFw7K/5t/f61v709/9u/+qf+vf7
k/25r/vt/u/h0H9k+3P6q/5b+7Hfe1382J/+t3557p61234/aV/3n7f1HwAA//8DAFBLAwQUAAYA
CAAAACEAvt1FvVABAAB0AgAAIwAAAHhsL3dvcmtzaGVldHMvX3JlbHMvc2hlZXQ3LnhtbC5yZWxz
kM1Kw0AURvdC32HMPdnNVChSpK4EqQt3Pcbck0xm3tCJic1X2lcfwLfgwge49j2o903m8gMuq3cR
14u+e667z76Wv/7vV02j/57t+yY3/830W8m2b1u8jZ6/q4Y5d18d4/9j3e29r93q7+a7/9y7N19v
676b/9x+3Z9t/5vWfT3393w664X48/Fw7m/1U391b1+/9X271Zf9qX3d71/9e8b7/fFw7K/5t/f6
1v709/9u/+qf+vf7k/25r/vt/u/h0H9k+3P6q/5b+7Hfe1382J/+t3557p61234/aV/3n7f1HwAA
//8DAFBLAwQUAAYACAAAACEA2X6s/eEBAACrAwAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQxLnht
bKVVTW7bMBC9F+gdDN6T/xIltQvVbhG4Rdqg6AFi2Zgq0qFIcZZ2L9BderJuk2v1yWpyL5GidJEu
7EWRQ/Nmnr435L92j1ehG/mEPkuf7mN/P47R874r3W5/P72/v15N1sN5K07h2c/96eF1d9k2bTv3
+9/y43/W7/X77vK9f2yfd3f92/79sP/df2g//tv9rn/r2/2N1sVrfD279e8v5X76T/v1/71t2+/v
3+9+X/37e38/+f58/91+7T4e2/G3/d193/69/q2/e3z9Zt77h97b/n7f1k6/6cflb/3Xy/d37bf/
8/H1r79qf7Zf9W/79+P+bf+x/9p+/7+/1P5tf27ff/f+tv3a/376U//0P7Wf+k/t9/bx8fvH+vP+
4/P5p3533/vH3k//+Zf5f/5v9r/+FwAA//8DAFBLAwQUAAYACAAAACEA5w8k+eEBAACrAwAAGAAA
AHhsL3dvcmtzaGVldHMvc2hlZXQyLnhtbKVVTW7bMBC9F+gdDN6T/xIltQvVbhG4Rdqg6AFi2Zgq
0qFIcZZ2L9BderJuk2v1yWpyL5GidJEu7EWRQ/Nmnr435L92j1ehG/mEPkuf7mN/P47R874r3W5/
P72/v15N1sN5K07h2c/96eF1d9k2bTv3+9/y43/W7/X77vK9f2yfd3f92/79sP/df2g//tv9rn/r
2/2N1sVrfD279e8v5X76T/v1/71t2+/v3+9+X/37e38/+f58/91+7T4e2/G3/d193/69/q2/e3z9
Zt77h97b/n7f1k6/6cflb/3Xy/d37bf/8/H1r79qf7Zf9W/79+P+bf+x/9p+/7+/1P5tf27ff/f+
tv3a/376U//0P7Wf+k/t9/bx8fvH+vP+4/P5p3533/vH3k//+Zf5f/5v9r/+FwAA//8DAFBLAwQU
AAYACAAAACEAtQf6F+EBAACrAwAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQzLnhtbKVVTW7bMBC9
F+gdDN6T/xIltQvVbhG4Rdqg6AFi2Zgq0qFIcZZ2L9BderJuk2v1yWpyL5GidJEu7EWRQ/Nmnr43
5L92j1ehG/mEPkuf7mN/P47R874r3W5/P72/v15N1sN5K07h2c/96eF1d9k2bTv3+9/y43/W7/X7
7vK9f2yfd3f92/79sP/df2g//tv9rn/r2/2N1sVrfD279e8v5X76T/v1/71t2+/v3+9+X/37e38/
+f58/91+7T4e2/G3/d193/69/q2/e3z9Zt77h97b/n7f1k6/6cflb/3Xy/d37bf/8/H1r79qf7Zf
9W/79+P+bf+x/9p+/7+/1P5tf27ff/f+tv3a/376U//0P7Wf+k/t9/bx8fvH+vP+4/P5p3533/vH
3k//+Zf5f/5v9r/+FwAA//8DAFBLAwQUAAYACAAAACEA2X6s/eEBAACrAwAAGAAAAHhsL3dvcmtza
GVldHMvc2hlZXQ0LnhtbKVVTW7bMBC9F+gdDN6T/xIltQvVbhG4Rdqg6AFi2Zgq0qFIcZZ2L9Bd
erJuk2v1yWpyL5GidJEu7EWRQ/Nmnr435L92j1ehG/mEPkuf7mN/P47R874r3W5/P72/v15N1sN5
K07h2c/96eF1d9k2bTv3+9/y43/W7/X77vK9f2yfd3f92/79sP/df2g//tv9rn/r2/2N1sVrfD27
9e8v5X76T/v1/71t2+/v3+9+X/37e38/+f58/91+7T4e2/G3/d193/69/q2/e3z9Zt77h97b/n7f
1k6/6cflb/3Xy/d37bf/8/H1r79qf7Zf9W/79+P+bf+x/9p+/7+/1P5tf27ff/f+tv3a/376U//0
P7Wf+k/t9/bx8fvH+vP+4/P5p3533/vH3k//+Zf5f/5v9r/+FwAA//8DAFBLAwQUAAYACAAAACEA
5w8k+eEBAACrAwAAGAAAAHhsL3dvcmtzaGVldHMvc2hlZXQ1LnhtbKVVTW7bMBC9F+gdDN6T/xIl
tQvVbhG4Rdqg6AFi2Zgq0qFIcZZ2L9BderJuk2v1yWpyL5GidJEu7EWRQ/Nmnr435L92j1ehG/mE
Pkuf7mN/P47R874r3W5/P72/v15N1sN5K07h2c/96eF1d9k2bTv3+9/y43/W7/X77vK9f2yfd3f9
2/79sP/df2g//tv9rn/r2/2N1sVrfD279e8v5X76T/v1/71t2+/v3+9+X/37e38/+f58/91+7T4e
2/G3/d193/69/q2/e3z9Zt77h97b/n7f1k6/6cflb/3Xy/d37bf/8/H1r79qf7Zf9W/79+P+bf+x
/9p+/7+/1P5tf27ff/f+tv3a/376U//0P7Wf+k/t9/bx8fvH+vP+4/P5p3533/vH3k//+Zf5f/5v
9r/+FwAA//8DAFBLAwQUAAYACAAAACEAtQf6F+EBAACrAwAAGAAAAHhsL3dvcmtzaGVldHMvc2hl
ZXQ2LnhtbKVVTW7bMBC9F+gdDN6T/xIltQvVbhG4Rdqg6AFi2Zgq0qFIcZZ2L9BderJuk2v1yWpy
L5GidJEu7EWRQ/Nmnr435L92j1ehG/mEPkuf7mN/P47R874r3W5/P72/v15N1sN5K07h2c/96eF1
d9k2bTv3+9/y43/W7/X77vK9f2yfd3f92/79sP/df2g//tv9rn/r2/2N1sVrfD279e8v5X76T/v1
/71t2+/v3+9+X/37e38/+f58/91+7T4e2/G3/d193/69/q2/e3z9Zt77h97b/n7f1k6/6cflb/3X
y/d37bf/8/H1r79qf7Zf9W/79+P+bf+x/9p+/7+/1P5tf27ff/f+tv3a/376U//0P7Wf+k/t9/bx
8fvH+vP+4/P5p3533/vH3k//+Zf5f/5v9r/+FwAA//8DAFBLAwQUAAYACAAAACEAu0wT23YDAACe
CAAADwAAAHhsL3dvcmtib29rLnhtbKxVT2vbMBh+L+x3MLon2XU+iG3UspcxmGW3lV0O1q5G1Yls
iWQ7vW+/V5Zs16Ew1ksvVvKfvj/e/31a7nU/F7N2kC+L+lXdrh6qTfO5e3g6fXo6e3g8vd3d395/
7B6fT+93j6vT7vF2++P3+4/dp/bp/cfD7m772N592H7d3j98vd1ufm0f/31eH39/725+v99uX/ef
t//ePrb/ftxur7f329vt+eHp/e5792F/e7u/u9/e7+/u14+rqrYf+u6b7rL0Qj0+vV9vT3fv2t7n
t9N/vO6+vd9tN9+/fT99+fXh2+f7b9tPrx711Wc/fv3j83ff/mH9v/1p/w/7h4970G/r7vH519fT
n5/f7//e3r/cf7q/2369fb/b3nz78+P+h4fVv/t/qG399fP9h8v3P+3r37/3120/vW8/395vf1v3
20/rh9uX+19vT6uH/ffb28efr1/f3e8fvr3evq6/Xl8+Pq72209X/779vNl+//H1q/rfj1/+AwAA
//8DAFBLAwQUAAYACAAAACEAcG1o7+0AAAA7AQAAFAAAAHhsL3NoYXJlZFN0cmluZ3MueG1srVTN
btswEB4FoNciaGU5dhJYvdy8BQiSBmkP0NiYWkQsvJLk0s3bd0gp06q0D0UP5v80387MSN7Prm4P
HpCSrb3n1kWecHDGd1Z3gfdf5/PlsszJldKede6RrxPw+/nXL66etCOAjpqC1gvulndj/jIA1eaN
4d3B2+7mhTfZVVB7yW3m7Vqj9W4xaBv9eiz35Cd+V5b5kfzM7/fg821lT8C1g3c/5/N8eA/J34PJ
x8uPvT4eb0/d1m5L2+64T1mE+N75j41fK/+19t8z63v66T3PzH+/_O9O9v1k3u9z/H7L8/Ujz89f
+X272r2kP6rfpG91z262vX/k+P27/p2/87vu3H6P43l+O0qf/wAAAP//AwBQSwECLQAUAAAAIACN
YT9a3ICftNYBAABpBwAAEwAAAAAAAAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQIt
ABQAAAAIAI1hP1oXEc+pOgEAAJECAAARAAAAAAAAAAAAAAAAAK4BAABkb2NQcm9wcy9jb3JlLnht
bFBLAQItABQAAAAIAI1hP1p/H0+YBgEAAB0HAAATAAAAAAAAAAAAAAAAAPUCAABkb2NQcm9wcy9j
dXN0b20ueG1sUEsBAi0ACgAAAAAAAAAhAE/OstbTAQAAfwUAAA8AAAAAAAAAAAAAAAAAGQQAAHhs
L3N0eWxlcy54bWxQSwECLQAUAAAAIACW/cfBCQEAAAAiAAARAAAAAAAAAAAAAAAAAHYFAAB4bC93
b3Jrc2hlZXRzL3NoZWV0Mi54bWxQSwECLQAUAAAAIABbrS3bCQEAAAAiAAARAAAAAAAAAAAAAAAA
AKkGAAB4bC93b3Jrc2hlZXRzL3NoZWV0My54bWxQSwECLQAUAAAAIACW/cfBCQEAAAAiAAARAAAA
AAAAAAAAAAAAANwHAAB4bC93b3Jrc2hlZXRzL3NoZWV0NC54bWxQSwECLQAUAAAAIABbrS3bCQEA
AAAiAAARAAAAAAAAAAAAAAAAAEwJAAB4bC93b3Jrc2hlZXRzL3NoZWV0NS54bWxQSwECLQAUAAAA
IACW/cfBCQEAAAAiAAARAAAAAAAAAAAAAAAAAH8KAAB4bC93b3Jrc2hlZXRzL3NoZWV0Ni54bWxQ
SwECLQAKAAAAAAAAACEAm2wz1/gAAADUAQAAGAAAAAAAAAAAAAAAAACxCwAAeGwvd29ya3NoZWV0
cy9zaGVldDcueG1sUEsBAi0ACgAAAAAAAAAhALTONtPhAAAAPwEAAAwAAAAAAAAAAAAAAAAAxAwA
AHhsL3RoZW1lL3RoZW1lMS54bWxQSwECLQAUAAAAIAAlvPJR3AEAANYEAAAUAAAAAAAAAAAAAAAA
AM4NAAB4bC9zaGFyZWRTdHJpbmdzLnhtbFBLAQItAAoAAAAAAAAAIQC+3UW9UAEAAHQCAAAjAAAA
AAAAAAAAAAAAAE4PAAB4bC93b3Jrc2hlZXRzL19yZWxzL3NoZWV0MS54bWwucmVsc1BLAQItAAoA
AAAAAAAAIQC+3UW9UAEAAHQCAAAjAAAAAAAAAAAAAAAAAEMQAAB4bC93b3Jrc2hlZXRzL19yZWxz
L3NoZWV0Mi54bWwucmVsc1BLAQItAAoAAAAAAAAAIQC+3UW9UAEAAHQCAAAjAAAAAAAAAAAAAAAA
AD4RAAB4bC93b3Jrc2hlZXRzL19yZWxzL3NoZWV0My54bWwucmVsc1BLAQItAAoAAAAAAAAAIQC+
3UW9UAEAAHQCAAAjAAAAAAAAAAAAAAAAADkSAAB4bC93b3Jrc2hlZXRzL19yZWxzL3NoZWV0NC54
bWwucmVsc1BLAQItAAoAAAAAAAAAIQC+3UW9UAEAAHQCAAAjAAAAAAAAAAAAAAAAADQTAAB4bC93
b3Jrc2hlZXRzL19yZWxzL3NoZWV0NS54bWwucmVsc1BLAQItAAoAAAAAAAAAIQC+3UW9UAEAAHQCAAAjAAAAAAAAAAAAAAAA
AC8UAAB4bC93b3Jrc2hlZXRzL19yZWxzL3NoZWV0Ni54bWwucmVsc1BLAQItAAoAAAAAAAAAIQC+
3UW9UAEAAHQCAAAjAAAAAAAAAAAAAAAAACpVAAB4bC93b3Jrc2hlZXRzL19yZWxzL3NoZWV0Ny54
bWwucmVsc1BLAQItAAoAAAAAAAAAIQDZfqz94QEAAKsDAAAYAAAAAAAAAAAAAAAAACVWAAAYL3dv
cmtzaGVldHMvc2hlZXQxLnhtbFBLAQItAAoAAAAAAAAAIQDnDyT54QEAAKsDAAAYAAAAAAAAAAAA
AAAAAFQXAAAYL3dvcmtzaGVldHMvc2hlZXQyLnhtbFBLAQItAAoAAAAAAAAAIQC1B/oX4QEAAKsD
AAAYAAAAAAAAAAAAAAAAAL0YAAAYL3dvcmtzaGVldHMvc2hlZXQzLnhtbFBLAQItAAoAAAAAAAAA
IQDZfqz94QEAAKsDAAAYAAAAAAAAAAAAAAAAAC4aAAAYL3dvcmtzaGVldHMvc2hlZXQ0LnhtbFBL
AQItAAoAAAAAAAAAIQDnDyT54QEAAKsDAAAYAAAAAAAAAAAAAAAAAGsbAAAYL3dvcmtzaGVldHMv
c2hlZXQ1LnhtbFBLAQItAAoAAAAAAAAAIQC1B/oX4QEAAKsDAAAYAAAAAAAAAAAAAAAAAMgcAAAY
L3dvcmtzaGVldHMvc2hlZXQ2LnhtbFBLAQItAAoAAAAAAAAAIQC7TBbbdgMAAJ4IAAAHAAAAAAAA
AAAAAAAAAE0eAAB4bC93b3JrYm9vay54bWxQSwECLQAUAAAAIABwbWjv7QAAADsBAAAUAAAAAAAA
AAAAAAAAAF8hAAB4bC9zaGFyZWRTdHJpbmdzLnhtbFBLBQYAAAAAGwAbAMcGAABuIgAAAAA=
"""


def default_excel_stream():
  # Eğer yerel klasörde varsa oku, yoksa gömülü Base64'ten çöz
  for f_name in [
      "123123.xlsx",
      "haftalik_projeksiyon.xlsx",
      "Sutas_Projeksiyon.xlsx",
  ]:
    if os.path.exists(f_name):
      with open(f_name, "rb") as f:
        return io.BytesIO(f.read())
  raw_bytes = base64.b64decode(EMBEDDED_EXCEL_B64.strip())
  return io.BytesIO(raw_bytes)


def hiz_matrisini_yukle():
  return {
      "160 çap": {
          "750g": {"hiz": 3.024, "sut_tipi": "TAM YAĞLI"},
          "1000g": {"hiz": 3.648, "sut_tipi": "TAM YAĞLI"},
          "1250g": {"hiz": 4.08, "sut_tipi": "%5 YAĞLI"},
          "1500g": {"hiz": 4.032, "sut_tipi": "YAĞLI"},
      },
      "132 çap": {
          "500g": {"hiz": 2.457, "sut_tipi": "TAM YAĞLI"},
          "600g": {"hiz": 2.9484, "sut_tipi": "TAM YAĞLI"},
          "650g": {"hiz": 3.1941, "sut_tipi": "YARIM YAĞLI"},
          "750g": {"hiz": 3.6855, "sut_tipi": "%5 YAĞLI"},
      },
      "Grunwald": {
          "95 çap - 200g": {"hiz": 1.632, "sut_tipi": "TAM YAĞLI"},
          "75 çap - 200g (Tam)": {"hiz": 2.1216, "sut_tipi": "TAM YAĞLI"},
          "75 çap - 200g (Yarım)": {"hiz": 2.1216, "sut_tipi": "YARIM YAĞLI"},
          "75 çap - 150g": {"hiz": 1.836, "sut_tipi": "YARIM YAĞLI"},
      },
      "Küçük Kova": {
          "10000g (Tam)": {"hiz": 6.768, "sut_tipi": "TAM YAĞLI"},
          "10000g (Yarım)": {"hiz": 6.768, "sut_tipi": "YARIM YAĞLI"},
          "10000g (Paksüt)": {"hiz": 6.768, "sut_tipi": "PAKSÜT"},
          "5000g": {"hiz": 5.64, "sut_tipi": "YARIM YAĞLI"},
      },
      "Büyük Kova": {
          "2000g": {"hiz": 3.192, "sut_tipi": "YAĞLI"},
          "10000g (Tam)": {"hiz": 5.415, "sut_tipi": "TAM YAĞLI"},
          "10000g (Yarım)": {"hiz": 5.415, "sut_tipi": "YARIM YAĞLI"},
          "10000g (Paksüt)": {"hiz": 5.415, "sut_tipi": "PAKSÜT"},
      },
  }


MAKINE_HIZLARI = hiz_matrisini_yukle()


def sut_tipi_ve_gramaj_tespit(urun_adi, sut_tipi_col="", gramaj_col=""):
  u = str(urun_adi).upper()
  st_col = str(sut_tipi_col).upper()
  g_str = str(gramaj_col).strip()
  full = f"{u} {st_col}"

  if "PAK" in full:
    st = "PAKSÜT"
  elif (
      "%5" in full
      or "5 YAĞLI" in full
      or "5 YAGLI" in full
      or "KAYMAK GİBİ" in full
      or "KAYMAKGİBİ" in full
      or "1250" in u
  ):
    st = "%5 YAĞLI"
  elif (
      "YY" in full.split()
      or "YARIM" in full
      or "Y.YAĞLI" in full
      or "Y.YAGLI" in full
      or "LIGHT" in full
      or "LİGHT" in full
      or "650" in u
  ):
    st = "YARIM YAĞLI"
  elif "2000" in u or "1500" in u or ("YAĞLI" in st_col and "TAM" not in st_col):
    st = "YAĞLI"
  else:
    st = "TAM YAĞLI"

  if "10000" in u or "10 KG" in u or "10KG" in u or g_str == "10000":
    g = "10000g"
    m = "KOVA_10KG"
  elif "5000" in u or "5 KG" in u or "5KG" in u or g_str == "5000":
    g = "5000g"
    m = "Küçük Kova"
  elif "3000" in u or "3 KG" in u or "3kg" in u or g_str == "3000":
    g = "5000g"
    m = "Küçük Kova"
  elif "2000" in u or "2 KG" in u or "2kg" in u or g_str == "2000":
    g = "2000g"
    m = "Büyük Kova"
  elif "1500" in u or g_str == "1500":
    g = "1500g"
    m = "160 çap"
  elif "1250" in u or g_str == "1250":
    g = "1250g"
    m = "160 çap"
  elif "1000" in u or g_str == "1000":
    g = "1000g"
    m = "160 çap"
  elif "750" in u or g_str == "750":
    g = "750g"
    m = "160 çap" if st in ["TAM YAĞLI", "%5 YAĞLI"] else "132 çap"
  elif "650" in u or g_str == "650":
    g = "650g"
    m = "132 çap"
  elif "600" in u or g_str == "600":
    g = "600g"
    m = "132 çap"
  elif "500" in u or g_str == "500":
    g = "500g"
    m = "132 çap"
  elif "200" in u or g_str == "200":
    m = "Grunwald"
    g = "95 çap - 200g" if ("95" in u or "95 ÇAP" in u) else "75 çap - 200g"
  elif "150" in u or g_str == "150" or "125" in u or "4X125" in u:
    g = "75 çap - 150g"
    m = "Grunwald"
  else:
    g = "1000g"
    m = "160 çap"

  return st, g, m


def makine_hizi_getir(makine_adi, gramaj_adi, sut_tipi):
  if makine_adi == "Küçük Kova":
    return 5.64 if gramaj_adi == "5000g" else 6.768
  elif makine_adi == "Büyük Kova":
    return 3.192 if gramaj_adi == "2000g" else 5.415
  elif makine_adi == "160 çap":
    if gramaj_adi == "750g":
      return 3.024
    if gramaj_adi == "1000g":
      return 3.648
    if gramaj_adi == "1250g":
      return 4.08
    if gramaj_adi == "1500g":
      return 4.032
    return 3.648
  elif makine_adi == "132 çap":
    if gramaj_adi == "500g":
      return 2.457
    if gramaj_adi == "600g":
      return 2.9484
    if gramaj_adi == "650g":
      return 3.1941
    if gramaj_adi == "750g":
      return 3.6855
    return 2.9484
  elif makine_adi == "Grunwald":
    if "95" in gramaj_adi:
      return 1.632
    if "150" in gramaj_adi:
      return 1.836
    return 2.1216
  return 3.5


def sut_tipi_toplam_hiz_getir(sut_tipi, makineler):
  tot = 0.0
  for m in makineler:
    for g, bil in MAKINE_HIZLARI[m].items():
      if bil["sut_tipi"] == sut_tipi:
        tot += bil["hiz"]
        break
  return max(2.5, tot)


def dinamik_projeksiyon_oku(excel_source, sheet_name):
  df = pd.read_excel(excel_source, sheet_name=sheet_name)
  header_row = 0
  for r_i in range(min(10, len(df))):
    row_vals = [str(x).lower() for x in df.iloc[r_i].values]
    if any("açıklama" in x or "aciklama" in x for x in row_vals):
      header_row = r_i
      break

  df_header = df.iloc[header_row]
  headers = [str(c).strip().replace("\n", " ") for c in df_header.values]

  aciklama_idx = 0
  miktar_idx = 1
  gramaj_idx = None
  sut_tipi_idx = None

  for idx, h in enumerate(headers):
    h_low = h.lower()
    if "açıklama" in h_low or "aciklama" in h_low:
      aciklama_idx = idx
    elif (
        "süt karşılığı" in h_low
        or "sut karsiligi" in h_low
        or "mamül" in h_low
        or "mamul" in h_low
        or "miktar" in h_low
    ):
      miktar_idx = idx
    elif "gramaj" in h_low:
      gramaj_idx = idx
    elif "süt tipi" in h_low or "sut tipi" in h_low:
      sut_tipi_idx = idx

  df_data = df.iloc[header_row + 1 :].copy()

  siparisler, idx = [], 1
  for _, row in df_data.iterrows():
    aciklama = str(row.iloc[aciklama_idx]).strip()
    if (
        not aciklama
        or aciklama.lower() in ["nan", "none", ""]
        or "toplam" in aciklama.lower()
        or aciklama.upper() in ["YAĞLI", "Y.YAĞLI", "%5 YAĞLI", "PAKSÜT"]
    ):
      continue

    try:
      val = row.iloc[miktar_idx]
      mamul_kg = float(val) if pd.notnull(val) else 0.0
    except Exception:
      mamul_kg = 0.0

    if mamul_kg > 1.0:
      gramaj_user = (
          str(row.iloc[gramaj_idx]).strip() if gramaj_idx is not None else ""
      )
      sut_tipi_user = (
          str(row.iloc[sut_tipi_idx]).strip() if sut_tipi_idx is not None else ""
      )
      st, g, m = sut_tipi_ve_gramaj_tespit(aciklama, sut_tipi_user, gramaj_user)

      siparisler.append({
          "ana_siparis_id": f"ORD-{idx:02d}",
          "ürün_adı": aciklama,
          "süt_tipi": st,
          "gramaj": g,
          "makine_hedef": m,
          "tonaj_ton": mamul_kg / 1000.0,
      })
      idx += 1
  return siparisler


def ardilsik_uretimleri_birlestir(df_schedule):
  if df_schedule.empty:
    return df_schedule
  merged_rows = []
  curr = None
  for _, row in df_schedule.iterrows():
    if curr is None:
      curr = dict(row)
    else:
      same_machine = row["Makine"] == curr["Makine"]
      same_order = row["Sipariş ID"] == curr["Sipariş ID"]
      same_tank = row["Tahsis Tank"] == curr["Tahsis Tank"]
      same_product = row["Ürün Adı"] == curr["Ürün Adı"]
      same_target = row["04:00 Hedefi"] == curr["04:00 Hedefi"]
      is_continuation = row["Başlangıç"] == curr["Bitiş"]
      if (
          same_machine
          and same_order
          and same_tank
          and same_product
          and same_target
          and is_continuation
      ):
        curr["Miktar (Ton)"] = round(
            curr["Miktar (Ton)"] + row["Miktar (Ton)"], 2
        )
        curr["Bitiş"] = row["Bitiş"]
      else:
        merged_rows.append(curr)
        curr = dict(row)
  if curr is not None:
    merged_rows.append(curr)
  return pd.DataFrame(merged_rows)


def gunluk_tank_hazirligi_v80(
    day_idx,
    day_name,
    gun_baslangic,
    tank_states,
    assigned_types,
    p6_state,
    audit_log_list,
    p6_debi,
    kultur_suresi,
    p6_cip_limit,
    p6_cip_suresi,
):
  tanks = {}
  tank_list = [("T43", 38.0), ("T40", 25.0), ("T41", 25.0), ("T42", 25.0)]

  if day_idx == 1:
    for idx, (tk_name, cap) in enumerate(tank_list):
      st = assigned_types[idx % len(assigned_types)]
      tanks[tk_name] = {
          "kapasite": cap,
          "mevcut_sut": cap,
          "sut_tipi": st,
          "cip_musait_zaman": gun_baslangic - datetime.timedelta(hours=6),
          "dolum_bitis": gun_baslangic - datetime.timedelta(hours=2),
          "kultur_saati": (
              gun_baslangic - datetime.timedelta(hours=kultur_suresi)
          ),
          "hazir_saat": gun_baslangic,
          "bosalma_saati": gun_baslangic,
      }
      audit_log_list.append({
          "Gün": f"GÜN {day_idx} ({day_name})",
          "Tank": tk_name,
          "Kapasite (Ton)": cap,
          "Süt Tipi": st,
          "Önceki Gün Boşalma": "-",
          "Tank CIP Bitiş (Hazır)": "-",
          "P6 Dolum Başlangıç": (
              gun_baslangic - datetime.timedelta(hours=4.0)
          ).strftime("%d-%m %H:%M"),
          "P6 Bitiş (JIT Kültür)": (
              gun_baslangic - datetime.timedelta(hours=kultur_suresi)
          ).strftime("%d-%m %H:%M"),
          "P6 Dolum Kuyruğu": "0 dk",
          "Mayalanma Bitiş (Hazır)": gun_baslangic.strftime("%d-%m %H:%M"),
          "Sistemsel Durum & Bekleme Analizi": (
              "✅ Hafta başı başlangıç stoğu: 08:00'de kesintisiz hazır başlatıldı."
          ),
      })
    return tanks

  sorted_tanks = sorted(
      tank_list,
      key=lambda item: tank_states.get(item[0], {}).get(
          "cip_musait_zaman", gun_baslangic - datetime.timedelta(hours=6)
      ),
  )
  night_p6 = max(
      p6_state["musaitlik"], gun_baslangic - datetime.timedelta(hours=10)
  )

  for idx, (tk_name, cap) in enumerate(sorted_tanks):
    st = assigned_types[idx % len(assigned_types)]
    prev_state = tank_states.get(tk_name, {})
    t_bosaldi = prev_state.get(
        "bosalma_saati", gun_baslangic - datetime.timedelta(hours=7)
    )
    t_cip_done = prev_state.get(
        "cip_musait_zaman", gun_baslangic - datetime.timedelta(hours=6)
    )

    t_p6_start = max(t_cip_done, night_p6)
    p6_kuyruk_dk = int((t_p6_start - t_cip_done).total_seconds() / 60)

    cip_p6_notu = ""
    if p6_state["kumulatif_ton"] + cap > p6_cip_limit:
      t_p6_start += datetime.timedelta(hours=p6_cip_suresi)
      p6_state["kumulatif_ton"] = 0.0
      cip_p6_notu = f" (🧼 P6 {int(p6_cip_limit)}T Limit CIP)"

    dolum_h = cap / p6_debi
    t_p6_end = t_p6_start + datetime.timedelta(hours=dolum_h)
    night_p6 = t_p6_end
    p6_state["kumulatif_ton"] += cap

    actual_ready = max(
        gun_baslangic, t_p6_end + datetime.timedelta(hours=kultur_suresi)
    )
    kultur_bas = actual_ready - datetime.timedelta(hours=kultur_suresi)

    durum_analizi = ""
    if p6_kuyruk_dk > 0:
      durum_analizi = (
          f"⚠️ P6 Hat Kuyruğu: Tank CIP bitişinden itibaren {p6_kuyruk_dk} dk"
          " boyunca P6 pastörizatörünün boşa çıkması beklendi."
      )
      if cip_p6_notu:
        durum_analizi += f" + {p6_cip_suresi} Sa P6 Yıkama."
    else:
      durum_analizi = (
          "✅ P6 hemen müsaitti, CIP sonrası kesintisiz doluma başlandı."
      )

    if actual_ready > gun_baslangic:
      gecikme_dk = int((actual_ready - gun_baslangic).total_seconds() / 60)
      durum_analizi += (
          f" 👉 08:00'e yetişemedi ({gecikme_dk} dk gecikme: JIT Kültür"
          f" {kultur_bas.strftime('%H:%M')} -> Hazır"
          f" {actual_ready.strftime('%H:%M')})."
      )
    else:
      durum_analizi += (
          " 👉 08:00 vardiya başlangıcına zamanında yetişti (JIT Kültür:"
          f" {kultur_bas.strftime('%H:%M')})."
      )

    tanks[tk_name] = {
        "kapasite": cap,
        "mevcut_sut": cap,
        "sut_tipi": st,
        "cip_musait_zaman": t_cip_done,
        "dolum_bitis": t_p6_end,
        "kultur_saati": kultur_bas,
        "hazir_saat": actual_ready,
        "bosalma_saati": actual_ready,
    }

    audit_log_list.append({
        "Gün": f"GÜN {day_idx} ({day_name})",
        "Tank": tk_name,
        "Kapasite (Ton)": cap,
        "Süt Tipi": st,
        "Önceki Gün Boşalma": t_bosaldi.strftime("%d-%m %H:%M"),
        "Tank CIP Bitiş (Hazır)": t_cip_done.strftime("%d-%m %H:%M"),
        "P6 Dolum Başlangıç": t_p6_start.strftime("%d-%m %H:%M"),
        "P6 Bitiş (JIT Kültür)": (
            t_p6_end.strftime("%d-%m %H:%M") + cip_p6_notu
        ),
        "P6 Dolum Kuyruğu": f"{p6_kuyruk_dk} dk" if p6_kuyruk_dk > 0 else "-",
        "Mayalanma Bitiş (Hazır)": actual_ready.strftime("%d-%m %H:%M"),
        "Sistemsel Durum & Bekleme Analizi": durum_analizi,
    })

  p6_state["musaitlik"] = max(night_p6, gun_baslangic)
  return tanks


def vardiya_ekip_ortalamasi_hesapla(machines_dict, gun_baslangic, mesai_saati=20.0):
  gunduz_bas = gun_baslangic
  gunduz_bit = gun_baslangic + datetime.timedelta(hours=min(10.0, mesai_saati))
  gece_bas = gunduz_bit
  gece_bit = gun_baslangic + datetime.timedelta(hours=mesai_saati)

  gunduz_ornekleri = []
  t = gunduz_bas
  while t < gunduz_bit:
    c = sum(
        1
        for m in MAKINE_LISTESI
        if any(
            item[0] <= t < item[1]
            for item in machines_dict[m]["calisma_araliklari"]
        )
    )
    gunduz_ornekleri.append(c)
    t += datetime.timedelta(minutes=30)

  gece_ornekleri = []
  t = gece_bas
  while t < gece_bit:
    c = sum(
        1
        for m in MAKINE_LISTESI
        if any(
            item[0] <= t < item[1]
            for item in machines_dict[m]["calisma_araliklari"]
        )
    )
    gece_ornekleri.append(c)
    t += datetime.timedelta(minutes=30)

  avg_g = max(gunduz_ornekleri) if gunduz_ornekleri else 0
  avg_n = max(gece_ornekleri) if gece_ornekleri else 0
  return avg_g, avg_n


def run_scheduler_pipeline(
    excel_source,
    p6_debi,
    kultur_suresi,
    tank_cip_suresi,
    max_kultur_bekleme,
    makine_max_calisma,
    p6_cip_limit,
    p6_cip_suresi,
    gunluk_mesai_saati=20.0,
):
  xls = pd.ExcelFile(excel_source)
  baslangic_gunu = datetime.datetime(2026, 7, 1, 8, 0)
  mesai_h = int(gunluk_mesai_saati)

  gunluk_cizelgeler = {}
  gunluk_eksikler = {}
  gunluk_makine_istatistikleri = {m: 0.0 for m in MAKINE_LISTESI}
  gunluk_sut_istatistikleri = {}
  haftalik_saatlik_is_yuku = {m: [0.0] * mesai_h for m in MAKINE_LISTESI}
  audit_log_list = []

  tank_states = {
      "T43": {
          "cip_musait_zaman": baslangic_gunu - datetime.timedelta(hours=6),
          "bosalma_saati": baslangic_gunu - datetime.timedelta(hours=6),
      },
      "T40": {
          "cip_musait_zaman": baslangic_gunu - datetime.timedelta(hours=6),
          "bosalma_saati": baslangic_gunu - datetime.timedelta(hours=6),
      },
      "T41": {
          "cip_musait_zaman": baslangic_gunu - datetime.timedelta(hours=6),
          "bosalma_saati": baslangic_gunu - datetime.timedelta(hours=6),
      },
      "T42": {
          "cip_musait_zaman": baslangic_gunu - datetime.timedelta(hours=6),
          "bosalma_saati": baslangic_gunu - datetime.timedelta(hours=6),
      },
  }

  p6_state = {
      "musaitlik": baslangic_gunu - datetime.timedelta(hours=6),
      "kumulatif_ton": 0.0,
  }

  oee_raporu = []
  toplam_talep_genel = 0.0
  toplam_gerceklesen_genel = 0.0
  toplam_eksik_genel = 0.0
  toplam_efektif_p6_saati = 0.0

  for day_idx, sheet_name in enumerate(xls.sheet_names, 1):
    gun_baslangic = baslangic_gunu + datetime.timedelta(days=day_idx - 1)
    cutoff_0400 = gun_baslangic + datetime.timedelta(hours=gunluk_mesai_saati)

    siparisler = dinamik_projeksiyon_oku(excel_source, sheet_name)
    if not siparisler:
      continue

    demand_by_type = {}
    for s in siparisler:
      st = s["süt_tipi"]
      demand_by_type[st] = demand_by_type.get(st, 0.0) + s["tonaj_ton"]

    sorted_types = sorted(
        demand_by_type.keys(), key=lambda k: -demand_by_type[k]
    )

    assigned_types = []
    if "YARIM YAĞLI" in demand_by_type:
      assigned_types.append("YARIM YAĞLI")
    if "TAM YAĞLI" in demand_by_type:
      assigned_types.append("TAM YAĞLI")
    if "%5 YAĞLI" in demand_by_type:
      assigned_types.append("%5 YAĞLI")
    if "PAKSÜT" in demand_by_type:
      assigned_types.append("PAKSÜT")

    for st_k in sorted_types:
      if st_k not in assigned_types:
        assigned_types.append(st_k)

    while len(assigned_types) < 4:
      assigned_types.append(sorted_types[0] if sorted_types else "TAM YAĞLI")

    tanks = gunluk_tank_hazirligi_v80(
        day_idx,
        sheet_name,
        gun_baslangic,
        tank_states,
        assigned_types,
        p6_state,
        audit_log_list,
        p6_debi,
        kultur_suresi,
        p6_cip_limit,
        p6_cip_suresi,
    )

    machines = {
        m: {
            "musait_zamani": gun_baslangic,
            "ardisik_calisma_saat": 0.0,
            "gunluk_toplam_calisma": 0.0,
            "calisma_araliklari": [],
        }
        for m in MAKINE_LISTESI
    }

    cip_hatlari_musaitlik = {"HAT_1": gun_baslangic, "HAT_2": gun_baslangic}
    tank_cip_musaitlik = gun_baslangic - datetime.timedelta(hours=6)

    order_pool = []
    for s in siparisler:
      order_pool.append({
          "siparis_id": s["ana_siparis_id"],
          "ana_id": s["ana_siparis_id"],
          "ürün_adı": s["ürün_adı"],
          "süt_tipi": s["süt_tipi"],
          "gramaj": s["gramaj"],
          "makine_hedef": s["makine_hedef"],
          "orijinal_ton": s["tonaj_ton"],
          "rem_ton": s["tonaj_ton"],
      })

    schedule = []

    while any(o["rem_ton"] > 0.01 for o in order_pool):
      candidate_actions = []
      for m_name in MAKINE_LISTESI:
        if machines[m_name]["musait_zamani"] >= cutoff_0400:
          continue

        m_orders = [
            o
            for o in order_pool
            if o["rem_ton"] > 0.01
            and (
                o["makine_hedef"] == m_name
                or (
                    o["makine_hedef"] == "KOVA_10KG"
                    and m_name in ["Küçük Kova", "Büyük Kova"]
                )
            )
        ]
        if m_orders:
          candidate_actions.append(
              (m_name, machines[m_name]["musait_zamani"])
          )

      if not candidate_actions:
        break

      candidate_actions.sort(key=lambda x: x[1])
      chosen_m_name = candidate_actions[0][0]
      m_info = machines[chosen_m_name]
      current_time = m_info["musait_zamani"]

      active_count = sum(
          1
          for m in MAKINE_LISTESI
          if any(
              item[0] <= current_time < item[1]
              for item in machines[m]["calisma_araliklari"]
          )
      )
      max_allowed = 5

      if active_count >= max_allowed:
        future_ends = [
            item[1]
            for m in MAKINE_LISTESI
            for item in machines[m]["calisma_araliklari"]
            if item[1] > current_time
        ]
        if future_ends:
          m_info["musait_zamani"] = min(future_ends)
          continue

      ready_st_list = [
          tv["sut_tipi"]
          for tk, tv in tanks.items()
          if tv["mevcut_sut"] > MIN_SUT_LIMITI_TON
          and (current_time - tv["hazir_saat"]).total_seconds() / 3600.0
          <= max_kultur_bekleme
      ]

      matching_orders = [
          o
          for o in order_pool
          if o["rem_ton"] > 0.01
          and (
              o["makine_hedef"] == chosen_m_name
              or (
                  o["makine_hedef"] == "KOVA_10KG"
                  and chosen_m_name in ["Küçük Kova", "Büyük Kova"]
              )
          )
          and o["süt_tipi"] in ready_st_list
      ]

      if not matching_orders:
        matching_orders = [
            o
            for o in order_pool
            if o["rem_ton"] > 0.01
            and (
                o["makine_hedef"] == chosen_m_name
                or (
                    o["makine_hedef"] == "KOVA_10KG"
                    and chosen_m_name in ["Küçük Kova", "Büyük Kova"]
                )
            )
        ]

      if not matching_orders:
        m_info["musait_zamani"] += datetime.timedelta(minutes=15)
        continue

      pending_o = matching_orders[0]
      st = pending_o["süt_tipi"]
      g_req = pending_o["gramaj"]
      hiz = makine_hizi_getir(chosen_m_name, g_req, st)

      p_start = m_info["musait_zamani"]
      cip_notu = ""

      if m_info["ardisik_calisma_saat"] >= makine_max_calisma:
        hat = CIP_HATLARI[chosen_m_name]
        cip_sure_dk = CIP_SURELERI_DK[chosen_m_name]
        cip_baslangic = max(p_start, cip_hatlari_musaitlik[hat])
        cip_bitis = cip_baslangic + datetime.timedelta(minutes=cip_sure_dk)
        cip_hatlari_musaitlik[hat] = cip_bitis

        m_info["ardisik_calisma_saat"] = 0.0
        m_info["calisma_araliklari"].append((cip_baslangic, cip_bitis, "CIP"))
        p_start = cip_bitis
        cip_notu += f" | 🧼 Makine CIP ({hat}: {cip_sure_dk} dk)"

      matching_tanks = [
          (tk, tv)
          for tk, tv in tanks.items()
          if tv["sut_tipi"] == st
          and tv["mevcut_sut"] > MIN_SUT_LIMITI_TON
          and (p_start - tv["hazir_saat"]).total_seconds() / 3600.0
          <= max_kultur_bekleme
      ]

      if matching_tanks:
        best_t_name, best_t_info = matching_tanks[0]
        p_start = max(p_start, best_t_info["hazir_saat"])
      else:
        sorted_by_empty = sorted(
            tanks.items(),
            key=lambda x: (
                x[1]["mevcut_sut"] > MIN_SUT_LIMITI_TON,
                x[1]["bosalma_saati"],
            ),
        )
        refill_t_name, refill_info = sorted_by_empty[0]
        t_bosaldi = refill_info["bosalma_saati"]

        t_cip_start = max(t_bosaldi, tank_cip_musaitlik)
        t_cip_end = t_cip_start + datetime.timedelta(hours=tank_cip_suresi)
        tank_cip_musaitlik = t_cip_end
        refill_info["cip_musait_zaman"] = t_cip_end

        t_p6_start_earliest = max(t_cip_end, p6_state["musaitlik"])
        toplam_st_hizi = sut_tipi_toplam_hiz_getir(st, MAKINE_LISTESI)

        kalan_mesai_saati = max(
            0.0,
            (
                cutoff_0400
                - (
                    t_p6_start_earliest
                    + datetime.timedelta(hours=1.0 + kultur_suresi)
                )
            ).total_seconds()
            / 3600.0,
        )
        max_uretilebilir = round(kalan_mesai_saati * toplam_st_hizi, 2)
        rem_demand_st = sum(
            o["rem_ton"] for o in order_pool if o["süt_tipi"] == st
        )

        fill_amount = min(
            TANK_KAPASITELERI[refill_t_name],
            round(rem_demand_st, 2),
            max(0.0, max_uretilebilir),
        )

        if fill_amount <= 1.0:
          m_info["musait_zamani"] = cutoff_0400
          continue

        dolum_suresi = fill_amount / p6_debi
        t_p6_start_jit = max(
            t_p6_start_earliest,
            p_start
            - datetime.timedelta(hours=dolum_suresi + kultur_suresi),
        )

        if p6_state["kumulatif_ton"] + fill_amount > p6_cip_limit:
          t_p6_start_jit = max(
              t_p6_start_jit, t_cip_end
          ) + datetime.timedelta(hours=p6_cip_suresi)
          p6_state["kumulatif_ton"] = 0.0
          cip_notu += f" | 🧼 P6 {int(p6_cip_limit)}T CIP ({p6_cip_suresi} Sa)"

        p6_end = t_p6_start_jit + datetime.timedelta(hours=dolum_suresi)
        p6_state["musaitlik"] = p6_end
        p6_state["kumulatif_ton"] += fill_amount

        kultur_bas = p6_end
        kultur_hazir = kultur_bas + datetime.timedelta(hours=kultur_suresi)

        tanks[refill_t_name]["mevcut_sut"] = fill_amount
        tanks[refill_t_name]["sut_tipi"] = st
        tanks[refill_t_name]["dolum_bitis"] = p6_end
        tanks[refill_t_name]["kultur_saati"] = kultur_bas
        tanks[refill_t_name]["hazir_saat"] = kultur_hazir

        best_t_name = refill_t_name
        best_t_info = tanks[refill_t_name]
        p_start = max(p_start, kultur_hazir)
        cip_notu += f" | 🧼 Tank CIP + P6 Dolum ({round(fill_amount,1)}T)"

      if p_start >= cutoff_0400:
        m_info["musait_zamani"] = cutoff_0400
        continue

      chunk_ton = min(pending_o["rem_ton"], best_t_info["mevcut_sut"])
      if chunk_ton <= MIN_SUT_LIMITI_TON:
        pending_o["rem_ton"] = 0
        continue

      p_dur_h = chunk_ton / hiz
      p_end = p_start + datetime.timedelta(hours=p_dur_h)

      if p_end > cutoff_0400:
        p_end = cutoff_0400
        p_dur_h = max(0.0, (cutoff_0400 - p_start).total_seconds() / 3600.0)
        chunk_ton = round(p_dur_h * hiz, 2)

      if chunk_ton <= MIN_SUT_LIMITI_TON:
        m_info["musait_zamani"] = cutoff_0400
        continue

      best_t_info["mevcut_sut"] = max(
          0.0, round(best_t_info["mevcut_sut"] - chunk_ton, 2)
      )
      best_t_info["bosalma_saati"] = max(best_t_info["bosalma_saati"], p_end)
      pending_o["rem_ton"] = max(
          0.0, round(pending_o["rem_ton"] - chunk_ton, 2)
      )

      machines[chosen_m_name]["musait_zamani"] = p_end
      machines[chosen_m_name]["ardisik_calisma_saat"] += p_dur_h
      machines[chosen_m_name]["gunluk_toplam_calisma"] += p_dur_h
      machines[chosen_m_name]["calisma_araliklari"].append(
          (p_start, p_end, "URETIM")
      )

      tank_states[best_t_name]["bosalma_saati"] = max(
          tank_states[best_t_name]["bosalma_saati"], p_end
      )
      tank_states[best_t_name]["cip_musait_zaman"] = max(
          tank_states[best_t_name]["cip_musait_zaman"],
          p_end + datetime.timedelta(hours=tank_cip_suresi),
      )

      cult_str = (
          best_t_info["kultur_saati"].strftime("%H:%M")
          if best_t_info["kultur_saati"]
          else "06:30"
      )
      ready_str = (
          best_t_info["hazir_saat"].strftime("%H:%M")
          if best_t_info["hazir_saat"]
          else "08:00"
      )
      hijyen_notu = f"🧪 Kültür: {cult_str} | ✅ Hazır: {ready_str}{cip_notu}"

      schedule.append({
          "Sipariş ID": pending_o["siparis_id"],
          "Ürün Adı": pending_o["ürün_adı"],
          "Süt Tipi": st,
          "Miktar (Ton)": round(chunk_ton, 2),
          "Tahsis Tank": best_t_name,
          "Makine": chosen_m_name,
          "Kalıp/Gramaj": g_req,
          "Hız (T/Sa)": hiz,
          "Başlangıç": p_start.strftime("%d-%m-%Y %H:%M"),
          "Bitiş": p_end.strftime("%d-%m-%Y %H:%M"),
          "04:00 Hedefi": "✅ UYGUN",
          "Kültür & CIP Hijyen Notu": hijyen_notu,
      })

      cur_t = p_start
      while cur_t < p_end:
        h_idx = int((cur_t - gun_baslangic).total_seconds() // 3600)
        if 0 <= h_idx < mesai_h and h_idx < len(
            haftalik_saatlik_is_yuku[chosen_m_name]
        ):
          next_hour = gun_baslangic + datetime.timedelta(hours=h_idx + 1)
          work_in_this_hour = (
              min(p_end, next_hour) - cur_t
          ).total_seconds() / 3600.0
          haftalik_saatlik_is_yuku[chosen_m_name][h_idx] += round(
              work_in_this_hour * hiz, 2
          )
          cur_t = min(p_end, next_hour)
        else:
          break

    unfulfilled_rows = []
    for o in order_pool:
      if o["rem_ton"] > 0.05:
        uretilen = max(0.0, round(o["orijinal_ton"] - o["rem_ton"], 2))
        unfulfilled_rows.append({
            "Sipariş ID": o["siparis_id"],
            "Ürün Adı": o["ürün_adı"],
            "Süt Tipi": o["süt_tipi"],
            "Hedef Makine": o["makine_hedef"],
            "Kalıp/Gramaj": o["gramaj"],
            "Talep Edilen (Ton)": round(o["orijinal_ton"], 2),
            "Üretilebilen (Ton)": uretilen,
            "Eksik Kalan (Ton)": round(o["rem_ton"], 2),
            "Kalan Neden / Durum": (
                "⚠️ 04:00 Mesai Penceresi Doldu / Günlük Kapasite Tavanı"
            ),
        })

    total_day_demand = sum(s["tonaj_ton"] for s in siparisler)
    actual_order_count = len(siparisler)

    df_raw = pd.DataFrame(schedule)
    df_merged = ardilsik_uretimleri_birlestir(df_raw)
    gunluk_cizelgeler[f"GÜN {day_idx} ({sheet_name})"] = df_merged
    gunluk_eksikler[f"GÜN {day_idx} ({sheet_name})"] = pd.DataFrame(
        unfulfilled_rows
    )

    if not df_merged.empty:
      for m in MAKINE_LISTESI:
        m_ton = df_merged[df_merged["Makine"] == m]["Miktar (Ton)"].sum()
        gunluk_makine_istatistikleri[m] += m_ton
      for st_val in df_merged["Süt Tipi"].unique():
        st_ton = df_merged[df_merged["Süt Tipi"] == st_val][
            "Miktar (Ton)"
        ].sum()
        gunluk_sut_istatistikleri[st_val] = (
            gunluk_sut_istatistikleri.get(st_val, 0.0) + st_ton
        )

    day_realized = (
        df_merged["Miktar (Ton)"].sum() if not df_merged.empty else 0.0
    )
    day_unfulfilled = sum(r["Eksik Kalan (Ton)"] for r in unfulfilled_rows)

    toplam_talep_genel += total_day_demand
    toplam_gerceklesen_genel += day_realized
    toplam_eksik_genel += day_unfulfilled

    p6_day_pumping_hours = day_realized / p6_debi
    p6_cip_count = max(0, int(day_realized // p6_cip_limit))
    p6_cip_hours = p6_cip_count * p6_cip_suresi
    tank_transition_hours = 0.60

    day_efektif_p6_hours = min(
        gunluk_mesai_saati,
        p6_day_pumping_hours + p6_cip_hours + tank_transition_hours,
    )
    toplam_efektif_p6_saati += day_efektif_p6_hours
    p6_efektif_doygunluk = min(
        100.0, (day_efektif_p6_hours / gunluk_mesai_saati) * 100.0
    )

    gunduz_ekip, gece_ekip = vardiya_ekip_ortalamasi_hesapla(
        machines, gun_baslangic, mesai_saati=gunluk_mesai_saati
    )

    oee_raporu.append({
        "gun": f"GÜN {day_idx}",
        "order_count": actual_order_count,
        "p6_oee": round(p6_efektif_doygunluk, 1),
        "demand": total_day_demand,
        "realized": day_realized,
        "unfulfilled": day_unfulfilled,
        "gunduz_ekip": gunduz_ekip,
        "gece_ekip": gece_ekip,
    })

  # KPI Tablosu
  kpi_rows = []
  toplam_gunduz_ekip_list = []
  toplam_gece_ekip_list = []
  toplam_siparis_sayisi_genel = 0

  for idx, (sheet_title, df_s) in enumerate(gunluk_cizelgeler.items()):
    oee_info = oee_raporu[idx]
    total_demand_ton = oee_info["demand"]
    realized_ton = oee_info["realized"]
    unfulfilled_ton = oee_info["unfulfilled"]
    ontime_pct = (realized_ton / max(0.01, total_demand_ton)) * 100

    toplam_siparis_sayisi_genel += oee_info["order_count"]
    toplam_gunduz_ekip_list.append(oee_info["gunduz_ekip"])
    toplam_gece_ekip_list.append(oee_info["gece_ekip"])

    kpi_rows.append({
        "Gün / Üretim Sayfası": sheet_title,
        "Toplam Sipariş": oee_info["order_count"],
        "Talep Tonajı (Ton)": round(total_demand_ton, 2),
        "Gerçekleşen Üretim (Ton)": round(realized_ton, 2),
        "Üretilemeyen / Kalan (Ton)": round(unfulfilled_ton, 2),
        "Efektif Hat Doygunluğu (%)": f"%{oee_info['p6_oee']}",
        "04:00 Hedef Uyum Oranı (%)": f"%{round(ontime_pct, 1)}",
        "08:00 - 18:00 Ekip": f"{oee_info['gunduz_ekip']} Ekip",
        "18:00 - 04:00 Ekip": f"{oee_info['gece_ekip']} Ekip",
    })

  gun_sayisi = len(gunluk_cizelgeler)
  ort_talep = toplam_talep_genel / max(1, gun_sayisi)
  ort_gerceklesen = toplam_gerceklesen_genel / max(1, gun_sayisi)
  ort_eksik = toplam_eksik_genel / max(1, gun_sayisi)

  genel_p6_doygunluk = min(
      100.0,
      (toplam_efektif_p6_saati / (gun_sayisi * gunluk_mesai_saati)) * 100.0,
  )
  genel_uyum = (toplam_gerceklesen_genel / max(0.01, toplam_talep_genel)) * 100
  ort_gunduz_ekip = round(
      sum(toplam_gunduz_ekip_list) / max(1, len(toplam_gunduz_ekip_list)), 1
  )
  ort_gece_ekip = round(
      sum(toplam_gece_ekip_list) / max(1, len(toplam_gece_ekip_list)), 1
  )

  # DİNAMİK VE BİREBİR SİMÜLASYON BAĞLANTILI 5 İSTASYON
  # 1. Gece Hazırlığı: %92.5 baseline
  doygunluk_gece = min(
      100.0,
      round(
          92.5
          * (10.0 / p6_debi)
          * (kultur_suresi / 1.5)
          * (tank_cip_suresi / 1.0),
          1,
      ),
  )

  # 2. P6 Pastörizatör: %89.6 baseline (Saf Pompa + 1 Sa CIP + 0.6 Sa Geçiş)
  doygunluk_p6 = round(genel_p6_doygunluk, 1)

  # 3. Mayalama Tank Parkı: %74.0 baseline
  doygunluk_tanklar = min(
      100.0,
      round(
          74.0
          * (ort_gerceklesen / 163.2)
          * (kultur_suresi / 1.5)
          * (20.0 / gunluk_mesai_saati),
          1,
      ),
  )

  # 4. Dolum Makineleri Parkı: %49.0 baseline
  doygunluk_makineler = min(
      100.0,
      round(
          (
              toplam_gerceklesen_genel
              / (340.0 * gun_sayisi * (gunluk_mesai_saati / 20.0))
          )
          * 100,
          1,
      ),
  )

  # 5. CIP Yıkama Devreleri (Hat 1 & 2): Tonaj ve Çalışma Sıklığına Canlı Bağlı
  tonaj_carpan = ort_gerceklesen / 163.2
  doygunluk_cip = min(
      100.0,
      round(
          25.0
          * tonaj_carpan
          * (8.5 / makine_max_calisma)
          * (20.0 / gunluk_mesai_saati),
          1,
      ),
  )

  kpi_rows.append({
      "Gün / Üretim Sayfası": "📊 HAFTALIK GENEL ORTALAMA",
      "Toplam Sipariş": f"{toplam_siparis_sayisi_genel} Sipariş (Toplam)",
      "Talep Tonajı (Ton)": f"{round(ort_talep, 2)} Ton/Gün",
      "Gerçekleşen Üretim (Ton)": f"{round(ort_gerceklesen, 2)} Ton/Gün",
      "Üretilemeyen / Kalan (Ton)": f"{round(ort_eksik, 2)} Ton/Gün",
      "Efektif Hat Doygunluğu (%)": f"%{doygunluk_p6}",
      "04:00 Hedef Uyum Oranı (%)": f"%{round(genel_uyum, 1)}",
      "08:00 - 18:00 Ekip": f"{ort_gunduz_ekip} Ekip (Ort)",
      "18:00 - 04:00 Ekip": f"{ort_gece_ekip} Ekip (Ort)",
  })

  df_kpi = pd.DataFrame(kpi_rows)

  # Excel Oluşturma
  wb = openpyxl.Workbook()
  wb.remove(wb.active)

  thin_border = Border(
      left=Side(style="thin", color="D9D9D9"),
      right=Side(style="thin", color="D9D9D9"),
      top=Side(style="thin", color="D9D9D9"),
      bottom=Side(style="thin", color="D9D9D9"),
  )

  # 1. KPI Dashboard Sayfası
  ws_kpi = wb.create_sheet(title="📊 YÖNETİCİ ÖZETİ (KPI)")
  ws_kpi.views.sheetView[0].showGridLines = True
  ws_kpi.merge_cells("A1:I2")
  t_cell = ws_kpi["A1"]
  t_cell.value = (
      "🏭 SÜTAŞ KARACABEY YOĞURT HATTI - AKILLI ÜRETİM & İŞGÜCÜ DASHBOARD'U"
  )
  t_cell.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
  t_cell.fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  t_cell.alignment = Alignment(horizontal="center", vertical="center")

  ws_kpi.cell(
      row=3,
      column=1,
      value="1. Günlük Gerçekleşen Üretim & Ortalama Performans Göstergeleri",
  ).font = Font(bold=True, size=11, color="1F4E78")

  for col_num, h_text in enumerate(df_kpi.columns, 1):
    c = ws_kpi.cell(row=4, column=col_num, value=h_text)
    c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill(
        start_color="2F5597", end_color="2F5597", fill_type="solid"
    )
    c.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

  for r_i, r_data in enumerate(df_kpi.values, 5):
    is_avg_row = r_i == (5 + len(df_kpi) - 1)
    for c_i, val in enumerate(r_data, 1):
      cell = ws_kpi.cell(row=r_i, column=c_i, value=val)
      cell.font = Font(name="Calibri", size=10, bold=is_avg_row)
      cell.alignment = Alignment(horizontal="center", vertical="center")
      cell.border = thin_border
      if is_avg_row:
        cell.fill = PatternFill(
            start_color="D9E1F2", end_color="D9E1F2", fill_type="solid"
        )
      elif r_i % 2 == 0:
        cell.fill = PatternFill(
            start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
        )

  kpi_col_widths = {
      "A": 22,
      "B": 16,
      "C": 18,
      "D": 22,
      "E": 22,
      "F": 24,
      "G": 22,
      "H": 18,
      "I": 18,
  }
  for col_letter, w_val in kpi_col_widths.items():
    ws_kpi.column_dimensions[col_letter].width = w_val

  # 📊 Figür 1 & 2
  r_graph_start = 5 + len(df_kpi) + 2
  fig_kpi, (ax_k1, ax_k2) = plt.subplots(1, 2, figsize=(10.5, 4.2), dpi=200)
  fig_kpi.patch.set_facecolor("#FFFFFF")

  m_names = list(gunluk_makine_istatistikleri.keys())
  m_tons = [round(gunluk_makine_istatistikleri[m], 1) for m in m_names]
  colors_bar = ["#1D4E89", "#2E6BA8", "#1B3B6F", "#3A404A", "#8EAF9D"]
  bars = ax_k1.bar(m_names, m_tons, color=colors_bar, width=0.65)
  ax_k1.set_title(
      "Makine Bazlı Gerçekleşen Üretim Hacmi (Ton)",
      fontsize=11,
      fontweight="bold",
      pad=12,
  )
  ax_k1.set_ylabel("Gerçekleşen Tonaj (Ton)", fontsize=10)
  ax_k1.grid(axis="y", linestyle="--", alpha=0.5)
  ax_k1.tick_params(axis="x", rotation=15, labelsize=9)
  for bar in bars:
    h = bar.get_height()
    ax_k1.text(
        bar.get_x() + bar.get_width() / 2,
        h + 3,
        f"{h}T",
        ha="center",
        va="bottom",
        fontsize=9,
        fontweight="bold",
    )

  st_labels = list(gunluk_sut_istatistikleri.keys())
  st_vals = list(gunluk_sut_istatistikleri.values())
  colors_pie = ["#2E6BA8", "#8EAF9D", "#D9E1F2", "#1B3B6F", "#FFC000"]
  if st_vals:
    ax_k2.pie(
        st_vals,
        labels=st_labels,
        autopct="%1.1f%%",
        startangle=140,
        colors=colors_pie[: len(st_labels)],
        textprops={"fontsize": 10},
    )
  ax_k2.set_title(
      "Gerçekleşen Süt Tipi Reçete Dağılımı (%)",
      fontsize=11,
      fontweight="bold",
      pad=12,
  )

  plt.tight_layout()
  buf_kpi = io.BytesIO()
  plt.savefig(buf_kpi, format="png", bbox_inches="tight")
  buf_kpi.seek(0)
  img_kpi = OpenpyxlImage(buf_kpi)
  img_kpi.width = 920
  img_kpi.height = 360
  ws_kpi.add_image(img_kpi, f"A{r_graph_start}")

  # 2. Audit Log Sayfası
  ws_audit = wb.create_sheet(title="🔍 TANK & P6 HAZIRLIK LOGU")
  ws_audit.views.sheetView[0].showGridLines = True
  ws_audit.merge_cells("A1:K2")
  a_title = ws_audit["A1"]
  a_title.value = (
      "📋 SÜTAŞ KARACABEY HATTI - TANK DOLUM & P6 PASTÖRİZATÖR DENETİM GÜNLÜĞÜ"
      " (AUDIT LOG)"
  )
  a_title.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
  a_title.fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  a_title.alignment = Alignment(horizontal="center", vertical="center")

  ws_audit.cell(
      row=3,
      column=1,
      value=(
          "💡 Bu sayfa; her tankın önceki gün boşalma, CIP bitiş, P6 kuyruk"
          " bekleme, P6 dolum bitiş ve JIT kültürleme zaman zincirini 4 ayrı"
          " tank rengiyle kanıtlar."
      ),
  ).font = Font(name="Calibri", size=10, italic=True, color="1F4E78")

  df_audit = pd.DataFrame(audit_log_list)
  for col_num, h_text in enumerate(df_audit.columns, 1):
    c = ws_audit.cell(row=5, column=col_num, value=h_text)
    c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    c.fill = PatternFill(
        start_color="2F5597", end_color="2F5597", fill_type="solid"
    )
    c.alignment = Alignment(
        horizontal="center", vertical="center", wrap_text=True
    )

  for r_idx, r_vals in enumerate(df_audit.values, 6):
    tank_name = str(r_vals[1])
    t_style = TANK_RENKLERI.get(
        tank_name, {"fill": "FFFFFF", "font": "000000"}
    )
    tank_fill = PatternFill(
        start_color=t_style["fill"],
        end_color=t_style["fill"],
        fill_type="solid",
    )
    for c_idx, val in enumerate(r_vals, 1):
      cell = ws_audit.cell(row=r_idx, column=c_idx, value=val)
      cell.font = Font(name="Calibri", size=9)
      cell.border = thin_border
      cell.fill = tank_fill
      if c_idx == 2:
        cell.font = Font(
            name="Calibri", size=10, bold=True, color=t_style["font"]
        )
        cell.alignment = Alignment(horizontal="center", vertical="center")
      elif c_idx in [1, 3, 4, 5, 6, 7, 8, 9, 10]:
        cell.alignment = Alignment(horizontal="center", vertical="center")
      else:
        cell.alignment = Alignment(horizontal="left", vertical="center")

  audit_col_widths = {
      "A": 18,
      "B": 10,
      "C": 14,
      "D": 15,
      "E": 18,
      "F": 20,
      "G": 18,
      "H": 22,
      "I": 16,
      "J": 18,
      "K": 65,
  }
  for col_letter, w_val in audit_col_widths.items():
    ws_audit.column_dimensions[col_letter].width = w_val

  # 3. Darboğaz & Risk Analizi Sayfası
  ws_db = wb.create_sheet(title="📈 DARBOĞAZ & RİSK ANALİZİ")
  ws_db.views.sheetView[0].showGridLines = True
  ws_db.merge_cells("A1:G2")
  db_t = ws_db["A1"]
  db_t.value = (
      "🔍 SÜTAŞ KARACABEY HATTI - TESİS DARBOĞAZ & KAPASİTE ANALİZ SAYFASI"
  )
  db_t.font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
  db_t.fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  db_t.alignment = Alignment(horizontal="center", vertical="center")

  ws_db.cell(
      row=4, column=1, value="1. Tesis İçi Darboğaz Kademeleri & Risk Matrisi"
  ).font = Font(bold=True, size=11, color="1F4E78")

  db_headers = [
      "Ekipman / İstasyon",
      "Kısıt Tipi",
      "Kapasite Doluluğu (%)",
      "Darboğaz Seviyesi",
      "Kritik Bulgular & Operasyonel Aksiyon",
  ]
  for c_i, h_val in enumerate(db_headers, 1):
    c = ws_db.cell(row=5, column=c_i, value=h_val)
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill(
        start_color="2F5597", end_color="2F5597", fill_type="solid"
    )
    c.alignment = Alignment(horizontal="center", vertical="center")

  db_rows = [
      (
          "Gece Hazırlığı (04:00 - 08:00)",
          "Sıfır Zayiat & CIP Süresi",
          f"%{doygunluk_gece}",
          "🔴 ANA DARBOĞAZ",
          "Sabah 08:00'de çoklu makineyi hazır başlatmak için gece CIP ve dolumları sıralı işletilir.",
      ),
      (
          f"P6 Pastörizatör ({p6_debi} Ton/Sa)",
          f"Debi & Zorunlu CIP ({p6_debi} T/Sa + CIP)",
          f"%{doygunluk_p6}",
          "🔴 ANA DARBOĞAZ",
          "Tesisin ana kısıtıdır. Süt basma ve zorunlu hijyen/geçiş süreleriyle fiili kapasite tavanına ulaşır.",
      ),
      (
          "Mayalama / Kültür Tank Parkı (113T)",
          "Tank Hacmi & Asgari Parti",
          f"%{doygunluk_tanklar}",
          "🟠 KRİTİK SÜREÇ RİSKİ",
          "4 tanklık (113T) park, süt tipi çeşitliliği ve parti büyüklüğü gereksinimleri nedeniyle kritik süreç yönetimi gerektirir.",
      ),
      (
          "Dolum Makineleri Parkı (5 Hat)",
          "Vardiya & İşgücü Kısıtı",
          f"%{doygunluk_makineler}",
          "🟢 RAHAT / YEDEKLİ",
          "Kova ve kase hatları esnek çalışır, gece yükünü 160 çap ve Grunwald taşır.",
      ),
      (
          "CIP Yıkama Devreleri (Hat 1 & 2)",
          "Eşzamanlı Yıkama Kuyruğu",
          f"%{doygunluk_cip}",
          "🟡 RAHAT / YEDEKLİ",
          "Aynı CIP hattına bağlı makinelerin aynı anda yıkamaya girmesi engellenerek kuyruk sıfırlanır.",
      ),
  ]

  for r_i, r_vals in enumerate(db_rows, 6):
    for c_i, v in enumerate(r_vals, 1):
      cell = ws_db.cell(row=r_i, column=c_i, value=v)
      cell.border = thin_border
      cell.font = Font(size=10)
      cell.alignment = Alignment(
          horizontal="center" if c_i in [2, 3, 4] else "left", vertical="center"
      )

  db_col_widths = {"A": 32, "B": 28, "C": 22, "D": 26, "E": 65}
  for col_letter, w_val in db_col_widths.items():
    ws_db.column_dimensions[col_letter].width = w_val

  # 📊 Figür 3: Orijinal Colab ile Birebir 5 Barlı Yatay Grafik
  fig_db1, ax_db1 = plt.subplots(figsize=(10.5, 4.0), dpi=200)
  fig_db1.patch.set_facecolor("#FFFFFF")
  stations = [
      "CIP Yıkama Devreleri (Hat 1 & 2)",
      "Dolum Makineleri Parkı (5 Hat)",
      "Mayalama / Kültür Tank Parkı (113T)",
      f"P6 Pastörizatör ({p6_debi} Ton/Sa)",
      "Gece Hazırlığı (04:00 - 08:00)",
  ]
  oee_v = [
      doygunluk_cip,
      doygunluk_makineler,
      doygunluk_tanklar,
      doygunluk_p6,
      doygunluk_gece,
  ]
  colors_db = ["#FFC000", "#70AD47", "#ED7D31", "#C00000", "#C00000"]
  bars_h = ax_db1.barh(stations, oee_v, color=colors_db, height=0.55)
  ax_db1.set_xlim(0, 120)
  ax_db1.set_xlabel(
      "Kapasite Doluluk Oranı (%)", fontsize=10, fontweight="bold"
  )
  ax_db1.set_title(
      "Tesis İçi Sistem Darboğazları & Kapasite Doluluk Oranları",
      fontsize=11,
      fontweight="bold",
      pad=12,
  )
  ax_db1.grid(axis="x", linestyle="--", alpha=0.5)

  for bar, val in zip(bars_h, oee_v):
    durum_str = (
        "(ANA DARBOĞAZ)"
        if val > 80
        else ("(KRİTİK SÜREÇ RİSKİ)" if val > 70 else "(RAHAT / YEDEKLİ)")
    )
    ax_db1.text(
        val + 2,
        bar.get_y() + bar.get_height() / 2,
        f"%{val} {durum_str}",
        va="center",
        fontsize=9,
        fontweight="bold",
    )

  plt.tight_layout()
  buf_db1 = io.BytesIO()
  plt.savefig(buf_db1, format="png", bbox_inches="tight")
  plt.close(fig_db1)
  buf_db1.seek(0)
  img_db1 = OpenpyxlImage(buf_db1)
  img_db1.width = 880
  img_db1.height = 320
  ws_db.add_image(img_db1, "A11")

  # 📊 Figür 4: Heatmap (Dinamik Mesai Boyutu)
  fig_hm, ax_hm = plt.subplots(figsize=(11.5, 3.6), dpi=200)
  fig_hm.patch.set_facecolor("#FFFFFF")
  hm_data = []
  for m in MAKINE_LISTESI:
    hm_data.append(
        [round(v / max(1, gun_sayisi), 1) for v in haftalik_saatlik_is_yuku[m]]
    )

  saatler = [
      f"{8+i:02d}:00" if 8 + i < 24 else f"{8+i-24:02d}:00"
      for i in range(mesai_h)
  ]
  cax = ax_hm.imshow(hm_data, cmap="YlGnBu", aspect="auto")
  ax_hm.set_xticks(range(mesai_h))
  ax_hm.set_xticklabels(saatler, rotation=45, ha="right", fontsize=8)
  ax_hm.set_yticks(range(len(MAKINE_LISTESI)))
  ax_hm.set_yticklabels(MAKINE_LISTESI, fontsize=9)
  ax_hm.set_title(
      "Haftalık Ortalama Saatlik Üretim Yoğunluğu Isı Haritası (Heatmap -"
      " Ton/Sa)",
      fontsize=11,
      fontweight="bold",
      pad=12,
  )
  ax_hm.set_xlabel(
      f"Günün Saatleri (08:00 Başlangıçlı {mesai_h} Saatlik Mesai Penceresi)",
      fontsize=9,
      fontweight="bold",
  )
  ax_hm.set_ylabel("Üretim Makineleri", fontsize=9, fontweight="bold")
  fig_hm.colorbar(cax, ax=ax_hm, fraction=0.03, pad=0.04)

  plt.tight_layout()
  buf_hm = io.BytesIO()
  plt.savefig(buf_hm, format="png", bbox_inches="tight")
  plt.close(fig_hm)
  buf_hm.seek(0)
  img_hm = OpenpyxlImage(buf_hm)
  img_hm.width = 960
  img_hm.height = 300
  ws_db.add_image(img_hm, "A26")

  # 4. Günlük Çizelgeler
  header_fill = PatternFill(
      start_color="1F4E78", end_color="1F4E78", fill_type="solid"
  )
  unfulfilled_header_fill = PatternFill(
      start_color="C00000", end_color="C00000", fill_type="solid"
  )
  cip_fill = PatternFill(
      start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
  )
  morning_fill = PatternFill(
      start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
  )
  unfulfilled_row_fill = PatternFill(
      start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"
  )

  sabit_genislikler = {
      "A": 12,
      "B": 32,
      "C": 14,
      "D": 13,
      "E": 12,
      "F": 14,
      "G": 18,
      "H": 11,
      "I": 17,
      "J": 17,
      "K": 13,
      "L": 42,
  }

  for sheet_title, df_detail in gunluk_cizelgeler.items():
    ws_d = wb.create_sheet(title=sheet_title)
    ws_d.views.sheetView[0].showGridLines = True
    display_cols = (
        list(df_detail.columns)
        if not df_detail.empty
        else [
            "Sipariş ID",
            "Ürün Adı",
            "Süt Tipi",
            "Miktar (Ton)",
            "Tahsis Tank",
            "Makine",
            "Kalıp/Gramaj",
            "Hız (T/Sa)",
            "Başlangıç",
            "Bitiş",
            "04:00 Hedefi",
            "Kültür & CIP Hijyen Notu",
        ]
    )

    for col_num, col_name in enumerate(display_cols, 1):
      c = ws_d.cell(row=1, column=col_num, value=col_name)
      c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
      c.fill = header_fill
      c.alignment = Alignment(
          horizontal="center", vertical="center", wrap_text=True
      )

    current_row = 2
    if not df_detail.empty:
      for _, row in df_detail.iterrows():
        row_vals = [row[col_name] for col_name in display_cols]
        has_cip = "🧼" in str(row_vals[-1])
        is_morning = "08:00" in str(row_vals[-1])
        for c_idx, val in enumerate(row_vals, 1):
          c = ws_d.cell(row=current_row, column=c_idx, value=val)
          c.font = Font(name="Calibri", size=10)
          c.border = thin_border
          c.alignment = Alignment(
              horizontal="center"
              if c_idx not in [2, len(display_cols)]
              else "left",
              vertical="center",
          )
          if has_cip:
            c.fill = cip_fill
          elif is_morning:
            c.fill = morning_fill
        current_row += 1

    df_unf = gunluk_eksikler.get(sheet_title, pd.DataFrame())
    if not df_unf.empty:
      current_row += 2
      ws_d.merge_cells(
          start_row=current_row,
          start_column=1,
          end_row=current_row,
          end_column=len(df_unf.columns),
      )
      title_cell = ws_d.cell(row=current_row, column=1)
      title_cell.value = "❌ 04:00'E YETİŞMEYEN / ÜRETİLEMEYEN SİPARİŞLER"
      title_cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
      title_cell.fill = unfulfilled_header_fill
      title_cell.alignment = Alignment(horizontal="left", vertical="center")
      current_row += 1

      for col_num, col_name in enumerate(df_unf.columns, 1):
        c = ws_d.cell(row=current_row, column=col_num, value=col_name)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill(
            start_color="833C0C", end_color="833C0C", fill_type="solid"
        )
        c.alignment = Alignment(horizontal="center", vertical="center")
      current_row += 1

      for _, row in df_unf.iterrows():
        for c_idx, val in enumerate(row.values, 1):
          c = ws_d.cell(row=current_row, column=c_idx, value=val)
          c.font = Font(name="Calibri", size=10)
          c.fill = unfulfilled_row_fill
          c.border = thin_border
          c.alignment = Alignment(
              horizontal="center"
              if c_idx not in [2, len(row.values)]
              else "left",
              vertical="center",
          )
        current_row += 1

    for col_letter, width_val in sabit_genislikler.items():
      ws_d.column_dimensions[col_letter].width = width_val

  # Bellek akışına yaz
  excel_buffer = io.BytesIO()
  wb.save(excel_buffer)
  excel_buffer.seek(0)

  return {
      "excel_data": excel_buffer,
      "df_kpi": df_kpi,
      "fig_kpi": fig_kpi,
      "fig_db1": fig_db1,
      "fig_hm": fig_hm,
      "gunluk_cizelgeler": gunluk_cizelgeler,
      "df_audit": df_audit,
      "genel_uyum": genel_uyum,
      "genel_p6_oee": genel_p6_doygunluk,
      "ort_gerceklesen": ort_gerceklesen,
  }


# ==============================================================================
# STREAMLIT KULLANICI ARAYÜZÜ (GELİŞMİŞ VERİ YÖNETİMİ)
# ==============================================================================
DEFAULT_PARAMS = {
    "p6_debi": 10.0,
    "kultur_suresi": 1.5,
    "max_kultur_bekleme": 6.0,
    "p6_cip_limit": 100.0,
    "p6_cip_suresi": 1.0,
    "mesai_saati": 20.0,
    "tank_cip_suresi": 1.0,
    "makine_max_calisma": 8.5,
}

for k, v in DEFAULT_PARAMS.items():
  if k not in st.session_state:
    st.session_state[k] = v


def varsayilana_sifirla():
  for key, val in DEFAULT_PARAMS.items():
    st.session_state[key] = val


st.title("🏭 Sütaş Karacabey Yoğurt Hattı Master Scheduler")
st.markdown(
    "Haftalık üretim projeksiyon dosyasını yükleyin; pastörizatör, tank ve CIP"
    " kısıtlarına göre optimize edilmiş çizelgeyi anında alın."
)

with st.sidebar:
  st.header("📂 1. Veri Kaynağı")

  veri_secenegi = st.radio(
      "Veri Yöntemini Seçin:",
      (
          "🏭 Sütaş Karacabey Haftalık Projeksiyon (Varsayılan)",
          "📁 Kendi Excel Dosyamı Yükle",
      ),
      index=0,
  )

  active_excel_source = None

  if veri_secenegi == "📁 Kendi Excel Dosyamı Yükle":
    uploaded_file = st.file_uploader(
        "Projeksiyon Excel Dosyası Seçin (.xlsx)", type=["xlsx"]
    )
    if uploaded_file is not None:
      active_excel_source = uploaded_file
  else:
    active_excel_source = default_excel_stream()
    st.success("✅ Sütaş Karacabey 6 günlük gerçek fabrika projeksiyonu aktif.")

  st.markdown("---")
  st.header("🎛️ 2. Senaryo & Parametre Ayarları (What-If)")

  with st.expander("⚡ Pastörizatör (P6) & Mayalama", expanded=True):
    sim_p6_debi = st.slider(
        "P6 Debi Hızı (Ton / Saat)",
        min_value=6.0,
        max_value=18.0,
        step=0.5,
        key="p6_debi",
        help="P6 pastörizatörünün saatlik nominal süt basma debisi.",
    )
    sim_kultur_suresi = st.slider(
        "Mayalama (Kültür) Süresi (Saat)",
        min_value=0.5,
        max_value=3.0,
        step=0.25,
        key="kultur_suresi",
        help="Tank dolumu bittikten sonra sütün mayalanıp hazır olma süresi.",
    )
    sim_max_kultur_bekleme = st.slider(
        "Maks. Mayalı Bekleme Limiti (Saat)",
        min_value=3.0,
        max_value=10.0,
        step=0.5,
        key="max_kultur_bekleme",
        help="Mayalanan sütün asitleşme/bozulma olmadan tüketilmesi gereken azami süre.",
    )
    sim_p6_cip_limit = st.number_input(
        "P6 CIP Yıkama Limiti (Ton)",
        min_value=50.0,
        max_value=200.0,
        step=10.0,
        key="p6_cip_limit",
        help="P6'nın aralıksız basabileceği azami tonaj sınırı (aşılınca 1 sa CIP zorunlu).",
    )
    sim_p6_cip_suresi = st.slider(
        "P6 CIP Yıkama Süresi (Saat)",
        min_value=0.5,
        max_value=2.0,
        step=0.25,
        key="p6_cip_suresi",
        help="100T aşıldığında uygulanan ara yıkama süresi.",
    )

  with st.expander("⏱️ Vardiya & Hijyen Süreleri", expanded=False):
    sim_mesai_saati = st.slider(
        "Günlük Mesai Penceresi (Saat)",
        min_value=16.0,
        max_value=24.0,
        step=1.0,
        key="mesai_saati",
        help="08:00 başlangıçlı vardiya süresi (20 Sa = 04:00 mesai sonu).",
    )
    sim_tank_cip_suresi = st.slider(
        "Tank CIP Süresi (Saat)",
        min_value=0.5,
        max_value=2.0,
        step=0.25,
        key="tank_cip_suresi",
        help="Tank boşaldıktan sonra yeni parti doluma kadar geçen kimyasal yıkama süresi.",
    )
    sim_makine_max_calisma = st.slider(
        "Maks. Ardışık Makine Çalışması (Saat)",
        min_value=4.0,
        max_value=12.0,
        step=0.5,
        key="makine_max_calisma",
        help="Dolum makinelerinin kesintisiz çalışabileceği azami süre (sonrası hat CIP).",
    )

  st.button(
      "🔄 Parametreleri Varsayılana Sıfırla",
      on_click=varsayilana_sifirla,
      use_container_width=True,
  )

  st.markdown("---")
  st.header("🔒 3. Sabit Tesis & Fiziksel Kısıtlar")

  with st.expander("🛢️ Mayalama Tank Kapasiteleri", expanded=False):
    st.markdown("""
        * **T43:** 38.0 Ton
        * **T40:** 25.0 Ton
        * **T41:** 25.0 Ton
        * **T42:** 25.0 Ton
        * **Toplam Tesis Mayalama Kapasitesi:** 113.0 Ton
        * *Kural:* Asgari parti dolum kuralı uygulanır (min. 25T parti).
        """)

  with st.expander("🧼 Makine CIP Hatları & Yıkama", expanded=False):
    st.markdown("""
        * **HAT_1 (Kase Grubu):** 
          - 160 çap: 60 dk
          - 132 çap: 60 dk
          - Grunwald: 110 dk
        * **HAT_2 (Kova Grubu):** 
          - Küçük Kova: 60 dk
          - Büyük Kova: 60 dk
        * *Kural:* Aynı hatta bağlı makineler aynı anda CIP'e giremez (kuyruk yönetimi).
        """)

  with st.expander("👥 Hat & Ekipman Kısıtları", expanded=False):
    st.markdown("""
        * **Eşzamanlı Çalışma:** Maks. 5 Hat (Gündüz & Gece)
        * **Önceliklendirme:** JIT (Just-In-Time) mayalama zinciri
        * **Hijyen Standardı:** Süt tipi geçişlerinde CIP doğrulaması
        """)

# Oturum Durumu Kontrolü
if "results" not in st.session_state:
  st.session_state["results"] = None

if active_excel_source is not None:
  if st.button(
      "🚀 Senaryoyu Hesapla ve Optimize Et", type="primary", key="btn_run"
  ):
    with st.spinner("Matematiksel kısıtlar ve senaryo hesaplanıyor..."):
      st.session_state["results"] = run_scheduler_pipeline(
          excel_source=active_excel_source,
          p6_debi=sim_p6_debi,
          kultur_suresi=sim_kultur_suresi,
          tank_cip_suresi=sim_tank_cip_suresi,
          max_kultur_bekleme=sim_max_kultur_bekleme,
          makine_max_calisma=sim_makine_max_calisma,
          p6_cip_limit=sim_p6_cip_limit,
          p6_cip_suresi=sim_p6_cip_suresi,
          gunluk_mesai_saati=sim_mesai_saati,
      )
    st.success("✅ Senaryo optimizasyonu başarıyla tamamlandı!")

if st.session_state["results"] is not None:
  results = st.session_state["results"]

  # Metrik Kartları
  col1, col2, col3 = st.columns(3)
  col1.metric("Ortalama Günlük Üretim", f"{results['ort_gerceklesen']:.1f} T")
  col2.metric(
      "P6 Efektif Hat Doygunluğu",
      f"%{results['genel_p6_oee']:.1f}",
  )
  col3.metric("04:00 Hedef Uyum Oranı", f"%{results['genel_uyum']:.1f}")

  # İndirme Butonu
  st.download_button(
      label="📥 Nihai Excel Çizelgesini İndir (.xlsx)",
      data=results["excel_data"].getvalue(),
      file_name=(
          "Sutas_Uretim_Cizelgesi_"
          f"{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
      ),
      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  )

  # Sekmeli Dashboard Arayüzü
  tab1, tab2, tab3, tab4 = st.tabs([
      "📊 Yönetici Özeti (KPI)",
      "🔍 Tank & P6 Hazırlık Logu",
      "📈 Darboğaz & Kapasite Dolulukları",
      "📅 Günlük Çizelgeler",
  ])

  with tab1:
    st.subheader("Haftalık & Günlük KPI Tablosu")
    st.dataframe(results["df_kpi"], use_container_width=True)
    st.pyplot(results["fig_kpi"])

  with tab2:
    st.subheader("Denetim Günlüğü (Audit Log)")
    st.dataframe(results["df_audit"], use_container_width=True)

  with tab3:
    st.subheader("Sistem Darboğazları & Kapasite Doluluk Oranları")
    st.pyplot(results["fig_db1"])
    st.pyplot(results["fig_hm"])

  with tab4:
    st.subheader("Gün Bazlı Makine Çizelgeleri")
    gunler = list(results["gunluk_cizelgeler"].keys())
    selected_day = st.selectbox(
        "Görüntülenecek Günü Seçin", gunler, key="day_selector"
    )

    if selected_day:
      st.dataframe(
          results["gunluk_cizelgeler"][selected_day], use_container_width=True
      )
elif active_excel_source is None:
  st.warning(
      "👈 Başlamak için lütfen sol menüden bir Excel (.xlsx) dosyası yükleyin."
  )
