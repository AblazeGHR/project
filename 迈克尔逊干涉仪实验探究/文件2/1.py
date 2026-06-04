import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Arc

# -------------------------- 画布设置
fig, ax = plt.subplots(figsize=(14, 8), dpi=120)
ax.set_xlim(-3, 12)
ax.set_ylim(-5, 5)
ax.axis('off')

# -------------------------- 几何参数
h_val = 0.6
L = 8
rho = 3.2
S1_x, S1_y = -2, 0
S2_x, S2_y = -2 + 2*h_val, 0
P_x, P_y = L, rho

# 光轴z轴（虚线）
ax.plot([-3, L+1], [0, 0], color="#888888", linestyle="--", lw=1.2)

# 画两点 S1 S2
ax.scatter(S1_x, S1_y, c="#992222", s=90, zorder=5)
ax.scatter(S2_x, S2_y, c="#ff9922", s=90, zorder=5)
ax.text(S1_x - 0.3, S1_y + 0.3, r"$S_1$", fontsize=14, color="#992222")
ax.text(S2_x - 0.3, S2_y + 0.3, r"$S_2$", fontsize=14, color="#ff7700")

# 观察屏竖线
ax.plot([L, L], [-4.5, 4.5], c="black", lw=1.4)
ax.text(L + 0.15, 4.3, "观察屏", fontsize=12)
ax.scatter(P_x, P_y, c="#2255dd", s=90, zorder=5)
ax.text(P_x + 0.2, P_y + 0.2, r"$P$", fontsize=13, color="#1133aa")

# 两条光线 S1→P(红), S2→P(橙)
ax.plot([S1_x, P_x], [S1_y, P_y], c="#992222", lw=1.8)
ax.plot([S2_x, P_x], [S2_y, P_y], c="#ff8822", lw=1.8)

# 标注2h线段
ax.plot([S1_x, S2_x], [-0.45, -0.45], c="black", lw=1.2)
ax.plot([S1_x, S1_x], [-0.45, -0.2], c="black", lw=1.2)
ax.plot([S2_x, S2_x], [-0.45, -0.2], c="black", lw=1.2)
ax.text((S1_x+S2_x)/2, -0.8, r"$2h$", ha="center", fontsize=13)

# L标注（S2到屏水平距离）
ax.plot([S2_x, L], [-0.55, -0.55], c="#555555", lw=1.2)
ax.plot([S2_x, S2_x], [-0.55, -0.3], c="#555555", lw=1.2)
ax.plot([L, L], [-0.55, -0.3], c="#555555", lw=1.2)
ax.text((S2_x+L)/2, -0.9, r"$L$", ha="center", fontsize=13, color="#444444")

# ρ标注（P到光轴竖直）
ax.plot([L+0.3, L+0.3], [0, P_y], c="#555555", lw=1.2)
ax.plot([L+0.1, L+0.3], [0, 0], c="#555555", lw=1.2)
ax.plot([L+0.1, L+0.3], [P_y, P_y], c="#555555", lw=1.2)
ax.text(L+0.55, P_y/2, r"$\rho$", va="center", fontsize=13, color="#444444")

# -------------------------- θ角绘制（S2处，光轴与S2P夹角）
arc_r = 1.1
theta_rad = np.arctan(rho/(L - S2_x))
arc = Arc((S2_x, S2_y), arc_r, arc_r, angle=0, theta1=0, theta2=np.rad2deg(theta_rad), color="darkblue", lw=1.5)
ax.add_patch(arc)
# θ文字摆放
txt_x = S2_x + arc_r*0.65*np.cos(theta_rad/2)
txt_y = S2_y + arc_r*0.65*np.sin(theta_rad/2)
ax.text(txt_x, txt_y, r"$\theta$", fontsize=14, color="darkblue")

# 右上角余弦公式
formula_text = r"$\cos\theta = \frac{L}{\sqrt{L^2+\rho^2}}$"
ax.text(L - 2.8, rho+0.4, formula_text, fontsize=15, color="#0044cc")

# 顶部标题与光程差公式(替换d→h)
ax.text(0, 4.2, "点光源迈克尔逊干涉仪：光程差几何示意图", ha="center", fontsize=15)
bbox_props = dict(boxstyle="round,pad=0.35", fc="#fef9b0", ec="#777777")
ax.text(0, 3.3, r"$\Delta = S_1P - S_2P \approx 2h\cos\theta$", ha="center", fontsize=16, bbox=bbox_props)

# 光轴标注
ax.text(L+0.7, -0.3, "光轴 $(z)$", fontsize=11, color="#333333")

plt.tight_layout()
plt.show()