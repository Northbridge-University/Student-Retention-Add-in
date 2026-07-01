import React from 'react';
import { User, Phone, X } from 'lucide-react';
import { formatPhoneNumber } from '../../utility/Conversion';

const cardStyles = `
@keyframes ecFadeIn {
  from { opacity: 0; transform: translateY(4px); }
  to   { opacity: 1; transform: translateY(0); }
}
@keyframes ecPulse {
  0%, 100% { opacity: 1; }
  50%      { opacity: 0.5; }
}
.ec-fade-in { animation: ecFadeIn 0.35s ease-out; }
.ec-skeleton-line { animation: ecPulse 1.2s ease-in-out infinite; background: #fecaca; border-radius: 6px; }
`;

// Skeleton placeholder shown while a new emergency contact is being saved.
export const EmergencyContactSkeleton = () => (
  <>
    <style>{cardStyles}</style>
    <div
      aria-hidden="true"
      style={{
        position: 'relative',
        border: '1px solid #fecaca',
        background: '#fef2f2',
        borderRadius: '0.75rem',
        padding: '0.75rem 0.875rem'
      }}
    >
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
        <div className="ec-skeleton-line" style={{ width: 16, height: 16, borderRadius: '9999px' }} />
        <div className="ec-skeleton-line" style={{ width: '45%', height: 12 }} />
        <div className="ec-skeleton-line" style={{ width: 56, height: 14, marginLeft: 'auto', borderRadius: '9999px' }} />
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <div className="ec-skeleton-line" style={{ width: 16, height: 16, borderRadius: '9999px' }} />
        <div className="ec-skeleton-line" style={{ width: '55%', height: 12 }} />
      </div>
    </div>
  </>
);

// Card that displays a single saved emergency contact.
// When `onRemove` is provided, a small delete button is shown.
const EmergencyContactCard = ({ contact, onRemove }) => {
  if (!contact) return null;

  const { name, number, relationship } = contact;

  return (
    <div
      className="ec-fade-in"
      style={{
        position: 'relative',
        border: '1px solid #fecaca',
        background: '#fef2f2',
        borderRadius: '0.75rem',
        padding: '0.75rem 0.875rem'
      }}
    >
      <style>{cardStyles}</style>
      {onRemove && (
        <button
          type="button"
          onClick={() => onRemove(contact)}
          aria-label={`Remove ${name || 'emergency contact'}`}
          title="Remove"
          style={{
            position: 'absolute',
            top: 6,
            right: 6,
            width: 22,
            height: 22,
            borderRadius: '9999px',
            border: 'none',
            background: 'transparent',
            color: '#b91c1c',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            cursor: 'pointer'
          }}
        >
          <X size={16} />
        </button>
      )}

      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6, paddingRight: 20 }}>
        <User size={16} color="#b91c1c" style={{ flexShrink: 0 }} />
        <span style={{ fontWeight: 700, color: '#1f2937', fontSize: 14 }}>
          {name || 'Unnamed contact'}
        </span>
        {relationship && (
          <span
            style={{
              marginLeft: 'auto',
              display: 'inline-flex',
              alignItems: 'center',
              fontSize: 11,
              fontWeight: 600,
              color: '#b91c1c',
              background: '#fee2e2',
              borderRadius: '9999px',
              padding: '2px 8px'
            }}
          >
            {relationship}
          </span>
        )}
      </div>

      {number && (
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <Phone size={16} color="#6b7280" style={{ flexShrink: 0 }} />
          <a
            href={`tel:${number}`}
            style={{ color: '#374151', fontSize: 14, textDecoration: 'none' }}
          >
            {formatPhoneNumber(number)}
          </a>
        </div>
      )}
    </div>
  );
};

export default EmergencyContactCard;
