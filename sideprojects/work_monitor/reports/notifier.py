"""Windows toast notifications for report completion."""

import logging

logger = logging.getLogger("work_monitor.notifier")


def send_toast(title: str, body: str, action_url: str = None):
    """Send a Windows toast notification.

    Args:
        title: Notification title.
        body: Notification body text.
        action_url: Optional file path to open on click.
    """
    try:
        from winotify import Notification, audio

        toast = Notification(
            app_id="Work Activity Monitor",
            title=title,
            msg=body[:200],  # Windows limits toast body length
            duration="long",
        )

        if action_url:
            toast.add_actions(label="Open Report", launch=action_url)

        toast.set_audio(audio.Default, loop=False)
        toast.show()
    except ImportError:
        logger.warning("winotify not installed, using console fallback")
        print(f"[NOTIFICATION] {title}: {body[:100]}")
    except Exception as e:
        logger.warning(f"Toast notification failed: {e}")
        print(f"[NOTIFICATION] {title}: {body[:100]}")
