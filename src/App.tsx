import { BrowserRouter, Routes, Route, Navigate } from 'react-router-dom'
import { AuthProvider } from '@/lib/auth'
import ProtectedRoute from '@/components/ProtectedRoute'
import AdminHome from '@/pages/AdminHome'
import NewProject from '@/pages/NewProject'
import ProjectAdmin from '@/pages/ProjectAdmin'
import ReviewEntry from '@/pages/ReviewEntry'
import ReviewMatrix from '@/pages/ReviewMatrix'
import ReviewInvite from '@/pages/ReviewInvite'
import ReviewSubmitted from '@/pages/ReviewSubmitted'

export default function App() {
  return (
    <BrowserRouter basename={import.meta.env.BASE_URL}>
      <AuthProvider>
        <Routes>
          <Route path="/" element={<Navigate to="/admin" replace />} />
          <Route path="/admin/login" element={<Navigate to="/admin" replace />} />
          <Route path="/auth/verify" element={<Navigate to="/admin" replace />} />

          <Route path="/admin" element={<ProtectedRoute><AdminHome /></ProtectedRoute>} />
          <Route path="/admin/projects/new" element={<ProtectedRoute><NewProject /></ProtectedRoute>} />
          <Route path="/admin/projects/:id" element={<ProtectedRoute><ProjectAdmin /></ProtectedRoute>} />

          <Route path="/review/invite/:inviteToken" element={<ReviewInvite />} />
          <Route path="/review/:reviewLinkToken" element={<ReviewEntry />} />
          <Route path="/review/:reviewLinkToken/matrix" element={<ReviewMatrix />} />
          <Route path="/review/:reviewLinkToken/submitted" element={<ReviewSubmitted />} />
        </Routes>
      </AuthProvider>
    </BrowserRouter>
  )
}
