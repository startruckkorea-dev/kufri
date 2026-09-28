/** 아직 정의되지 않은 화면. 화면 정의(리스트·피벗)가 정해지면 교체한다. */
export default function Placeholder({ item }) {
  return (
    <div className="card">
      <h2>{item.title}</h2>
      <div className="empty">
        준비 중입니다.
        <br />
        <span className="muted">이 화면에서 볼 리스트와 피벗 정의가 정해지면 구현합니다.</span>
      </div>
    </div>
  );
}

export function NotFound() {
  return (
    <div className="card">
      <h2>없는 화면</h2>
      <div className="empty">
        주소에 해당하는 화면이 없습니다. <a href="#/">첫 화면으로</a>
      </div>
    </div>
  );
}

export function LoginCard({ booting, onLogin }) {
  return (
    <div className="card">
      <div className="empty">
        <button className="btn primary" onClick={onLogin} disabled={booting}>
          Microsoft 계정으로 로그인
        </button>
        <br />
        <span className="muted">로그인한 계정에 부여된 SharePoint 권한 범위에서만 동작합니다.</span>
      </div>
    </div>
  );
}
