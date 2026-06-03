import matplotlib.pyplot as plt
import numpy as np
from pathlib import Path

plt.rcParams['font.sans-serif'] = ['SimHei', 'DejaVu Sans']
plt.rcParams['axes.unicode_minus'] = False

fig, ax = plt.subplots(figsize=(10, 6))
ax.set_aspect('equal')
ax.set_xlim(-4, 11)
ax.set_ylim(-1.5, 5.5)
ax.axis('off')
ax.set_title('点光源迈克尔逊干涉仪：光程差几何示意图', fontsize=14, fontweight='bold')

# 光轴
ax.axhline(y=0, color='gray', linestyle='--', linewidth=0.8, alpha=0.5)
ax.text(10.8, -0.2, '光轴 ($z$)', fontsize=11, ha='right', va='top', color='gray')

# 观察屏
ax.plot([10, 10], [-1.2, 5.5], color='black', linewidth=2)
ax.text(10.2, 5.5, '观察屏', fontsize=12, ha='left', va='bottom')

# 两个虚点光源
s2 = np.array([0., 0.])
s1 = np.array([-2.5, 0.])   # 间距 2d = 2.5 便于看清
ax.plot(*s2, 'o', color='darkorange', markersize=10, zorder=5)
ax.plot(*s1, 'o', color='darkred', markersize=10, zorder=5)
ax.text(s2[0]-0.2, s2[1]+0.3, '$S_2$', fontsize=13, ha='right', color='darkorange')
ax.text(s1[0]-0.2, s1[1]+0.3, '$S_1$', fontsize=13, ha='right', color='darkred')

# 标注 2d
ax.annotate('', xy=(s1[0], -0.5), xytext=(s2[0], -0.5),
            arrowprops=dict(arrowstyle='<->', color='black', lw=1.5))
ax.text(np.mean([s1[0], s2[0]]), -0.8, '$2d$', fontsize=14, ha='center')

# 观察点 P
L = 10.0
rho = 3.2
P = np.array([L, rho])
ax.plot(*P, 'o', color='blue', markersize=8, zorder=5)
ax.text(P[0]+0.2, P[1]+0.15, '$P$', fontsize=13, color='blue')

# 光线
ax.plot([s2[0], P[0]], [s2[1], P[1]], color='darkorange', linewidth=1.5)
ax.plot([s1[0], P[0]], [s1[1], P[1]], color='darkred', linewidth=1.5)

# 辅助线：P 到光轴垂足
ax.plot([P[0], P[0]], [0, P[1]], 'gray', linestyle=':', linewidth=0.8)
ax.text(P[0]+0.2, P[1]/2, '$\\rho$', fontsize=13, ha='left', va='center', color='gray')

# L 的标注
ax.annotate('', xy=(10, -0.1), xytext=(0, -0.1),
            arrowprops=dict(arrowstyle='<->', color='gray', lw=1.2))
ax.text(5, -0.4, '$L$', fontsize=13, ha='center', color='gray')

# 角度 θ
# 以 P 为顶点，水平向左的线与 S2-P 连线的夹角
theta = np.arctan2(rho, L)
arc_radius = 0.9
arc_angles = np.linspace(0, theta, 20)
ax.plot(P[0] - arc_radius * np.cos(arc_angles),
        P[1] - arc_radius * np.sin(arc_angles),
        'blue', lw=1.8)
ax.text(P[0]-1.3, P[1]-0.7, r'$\theta$', fontsize=15, color='blue')

# 标明 cosθ 关系
ax.text(P[0]-1.2, P[1]+0.5,
        r'$\cos\theta = \frac{L}{\sqrt{L^2+\rho^2}}$',
        fontsize=12, color='blue', ha='center')

# 最终光程差公式
ax.text(0.3, 5.0, r'$\mathbf{\Delta = S_1P - S_2P \approx 2d\cos\theta}$',
        fontsize=16, color='black', bbox=dict(boxstyle='round,pad=0.3',
                                              facecolor='lightyellow', edgecolor='gray'))

out_file = Path(__file__).parent / 'interference.png'
plt.tight_layout()
plt.savefig(out_file, dpi=300, bbox_inches='tight')
print(f"Saved figure to {out_file}")
plt.show()