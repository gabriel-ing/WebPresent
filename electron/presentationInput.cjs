const PRESENTATION_INPUT_ACTIONS = {
  next: 'next',
  previous: 'previous',
  exit: 'exit',
};

function getPresentationInputAction(input) {
  if (!input || input.type !== 'keyDown' || input.isAutoRepeat) {
    return null;
  }

  if (input.key === 'ArrowRight' && input.shift) {
    return PRESENTATION_INPUT_ACTIONS.next;
  }

  if (input.key === 'ArrowLeft' && input.shift) {
    return PRESENTATION_INPUT_ACTIONS.previous;
  }

  if (input.key === 'Escape') {
    return PRESENTATION_INPUT_ACTIONS.exit;
  }

  return null;
}

module.exports = {
  PRESENTATION_INPUT_ACTIONS,
  getPresentationInputAction,
};
