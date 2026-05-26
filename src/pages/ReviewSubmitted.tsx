import NavBar from '@/components/NavBar'

export default function ReviewSubmitted() {
  return (
    <div style={{ minHeight: '100vh', backgroundColor: '#F7F5F0' }}>
      <NavBar />
      <div style={{ display: 'flex', justifyContent: 'center', alignItems: 'center', minHeight: 'calc(100vh - 60px)', padding: '40px 16px' }}>
        <div style={{ textAlign: 'center', maxWidth: '480px', width: '100%' }}>
          {/* Checkmark icon */}
          <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '24px' }}>
            <div style={{
              width: '72px', height: '72px', borderRadius: '50%',
              backgroundColor: '#C8E6D4', display: 'flex', alignItems: 'center', justifyContent: 'center',
            }}>
              <svg width="36" height="36" viewBox="0 0 36 36" fill="none">
                <path d="M8 18l7 7L28 11" stroke="#00602B" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />
              </svg>
            </div>
          </div>

          <h1 style={{
            fontFamily: 'Montserrat, sans-serif', fontWeight: 700,
            fontSize: '24px', color: '#2C2C2C', marginBottom: '12px',
          }}>
            Review Submitted
          </h1>
          <p style={{
            fontFamily: 'Open Sans, sans-serif', fontSize: '15px',
            color: '#5C5C5C', lineHeight: 1.7, marginBottom: '0',
          }}>
            Your review has been successfully submitted. Thank you for your contribution — GIW Environmental Solutions will be in touch if any follow-up is required.
          </p>
        </div>
      </div>
    </div>
  )
}
