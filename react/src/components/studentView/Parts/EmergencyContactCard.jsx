import React from 'react';
import { User, Phone, X } from 'lucide-react';
import { formatPhoneNumber } from '../../utility/Conversion';

// Card that displays a single saved emergency contact.
// When `onRemove` is provided, a small delete button is shown.
const EmergencyContactCard = ({ contact, onRemove }) => {
  if (!contact) return null;

  const { name, number, relationship } = contact;

  return (
    <div
      style={{
        position: 'relative',
        border: '1px solid #fecaca',
        background: '#fef2f2',
        borderRadius: '0.75rem',
        padding: '0.75rem 0.875rem'
      }}
    >
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
